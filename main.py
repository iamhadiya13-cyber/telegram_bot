"""
Jivandeep Clinic - Telegram Appointment Bot
Railway-compatible | Python 3.13 compatible
"""

import asyncio
import logging
import os
from datetime import datetime, timedelta, time, date as date_type

import openpyxl
from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove, BotCommand, BotCommandScopeDefault, BotCommandScopeChat
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    ConversationHandler,
    filters,
    ContextTypes,
)

# ─────────────────────────────────────────────
# CONFIGURATION — use Railway environment variables
# ─────────────────────────────────────────────
BOT_TOKEN  = os.environ.get("BOT_TOKEN", "")
OWNER_ID   = int(os.environ.get("OWNER_ID", "0"))
EXCEL_FILE = "appointments.xlsx"
# ─────────────────────────────────────────────

WAITING_FOR_YES, ASK_NAME, ASK_AGE, ASK_REASON, ASK_MOBILE, ASK_DATE = range(6)

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)


# ══════════════════════════════════════════════
#  EXCEL HELPERS
# ══════════════════════════════════════════════

def initialize_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Appointments"
        ws.append(["Name", "Age", "Reason", "Mobile",
                   "Date", "Slot Time", "Booking Timestamp"])
        wb.save(EXCEL_FILE)


def save_appointment(data: dict):
    initialize_excel()
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    ws.append([
        data["name"],
        data["age"],
        data["reason"],
        data["mobile"],
        data["date"].strftime("%d/%m/%Y"),
        data["slot_time"],
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    ])
    try:
        wb.save(EXCEL_FILE)
    except PermissionError:
        raise PermissionError("appointments.xlsx is open. Please close it and try again.")


def get_booked_slots(desired_date: date_type) -> list:
    initialize_excel()
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    booked = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        try:
            booked_date = datetime.strptime(row[4], "%d/%m/%Y").date()
            if booked_date == desired_date:
                booked.append(row[5])
        except (ValueError, TypeError):
            continue
    return booked


# ══════════════════════════════════════════════
#  SLOT LOGIC
# ══════════════════════════════════════════════

def get_available_slot(desired_date: date_type):
    WORK_START  = time(9, 0)
    WORK_END    = time(22, 0)
    SUNDAY_END  = time(13, 0)
    LUNCH_START = time(13, 0)
    LUNCH_END   = time(16, 0)
    SLOT_MINS   = timedelta(minutes=30)

    is_sunday     = desired_date.weekday() == 6
    effective_end = SUNDAY_END if is_sunday else WORK_END

    booked_slots = get_booked_slots(desired_date)
    slot_start   = datetime.combine(desired_date, WORK_START)
    end_dt       = datetime.combine(desired_date, effective_end)

    while slot_start < end_dt:
        slot_end = slot_start + SLOT_MINS
        if slot_end > end_dt:
            break
        if not is_sunday and LUNCH_START <= slot_start.time() < LUNCH_END:
            slot_start = datetime.combine(desired_date, LUNCH_END)
            continue
        slot_str = f"{slot_start.strftime('%I:%M %p')} - {slot_end.strftime('%I:%M %p')}"
        if slot_str not in booked_slots:
            return slot_str
        slot_start += SLOT_MINS
    return None


def is_clinic_closed(desired_date: date_type):
    now = datetime.now()
    if desired_date.weekday() == 6:
        if desired_date == now.date() and now.time() >= time(13, 0):
            return "Jivandeep Clinic is closed on Sundays after 1 PM."
    return None


# ══════════════════════════════════════════════
#  CONVERSATION HANDLERS
# ══════════════════════════════════════════════

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    keyboard = [["YES"]]
    await update.message.reply_text(
        "🏥 *Welcome to Jivandeep Clinic!*\n\n"
        "Would you like to book an appointment slot?\n"
        "Reply *YES* to proceed.",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True),
    )
    return WAITING_FOR_YES


async def waiting_for_yes(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if update.message.text.strip().upper() == "YES":
        await update.message.reply_text(
            "👤 Please enter your *Full Name*:",
            parse_mode="Markdown",
            reply_markup=ReplyKeyboardRemove(),
        )
        return ASK_NAME
    keyboard = [["YES"]]
    await update.message.reply_text(
        "Please reply *YES* to proceed with booking. 😊",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True),
    )
    return WAITING_FOR_YES


async def ask_name(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    name = update.message.text.strip()
    if len(name) < 2:
        await update.message.reply_text("⚠️ Name too short. Please enter your full name:")
        return ASK_NAME
    context.user_data["name"] = name
    await update.message.reply_text("🔢 Please enter your *Age*:", parse_mode="Markdown")
    return ASK_AGE


async def ask_age(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    try:
        age = int(update.message.text.strip())
        if not (1 <= age <= 120):
            raise ValueError
        context.user_data["age"] = age
        await update.message.reply_text("📋 Please enter the *Reason for visit*:", parse_mode="Markdown")
        return ASK_REASON
    except ValueError:
        await update.message.reply_text("⚠️ Invalid age. Please enter a number between 1 and 120:")
        return ASK_AGE


async def ask_reason(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["reason"] = update.message.text.strip()
    await update.message.reply_text("📱 Please enter your *Mobile Number* (digits only):", parse_mode="Markdown")
    return ASK_MOBILE


async def ask_mobile(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    mobile  = update.message.text.strip()
    cleaned = mobile.lstrip("+")
    if not cleaned.isdigit() or not (7 <= len(cleaned) <= 15):
        await update.message.reply_text("⚠️ Invalid number. Please enter digits only (e.g. 9876543210):")
        return ASK_MOBILE
    context.user_data["mobile"] = mobile
    keyboard = [["Today"]]
    await update.message.reply_text(
        "📅 Would you like to book for *TODAY* or a *SPECIFIC DATE*?\n"
        "Reply *Today* or enter a date in *DD/MM/YYYY* format.",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True),
    )
    return ASK_DATE


async def ask_date(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    text  = update.message.text.strip()
    today = datetime.now().date()

    if text.lower() == "today":
        desired_date = today
    else:
        try:
            desired_date = datetime.strptime(text, "%d/%m/%Y").date()
        except ValueError:
            await update.message.reply_text(
                "⚠️ Invalid format. Reply *Today* or enter date as *DD/MM/YYYY* (e.g. 25/12/2026):",
                parse_mode="Markdown",
            )
            return ASK_DATE

    if desired_date < today:
        await update.message.reply_text(
            "⚠️ Cannot book for a past date. Please choose today or a future date (DD/MM/YYYY):",
            reply_markup=ReplyKeyboardRemove(),
        )
        return ASK_DATE

    closed_reason = is_clinic_closed(desired_date)
    if closed_reason:
        await update.message.reply_text(
            f"⚠️ {closed_reason}\nPlease choose another date (DD/MM/YYYY):",
            reply_markup=ReplyKeyboardRemove(),
        )
        return ASK_DATE

    available_slot = get_available_slot(desired_date)
    if not available_slot:
        await update.message.reply_text(
            "😔 No available slots for that date. Please choose another date (DD/MM/YYYY):",
            reply_markup=ReplyKeyboardRemove(),
        )
        return ASK_DATE

    context.user_data["date"]      = desired_date
    context.user_data["slot_time"] = available_slot

    try:
        save_appointment(context.user_data)
    except PermissionError:
        await update.message.reply_text(
            "⚠️ Could not save appointment. Please try again in a moment.",
            reply_markup=ReplyKeyboardRemove(),
        )
        return ConversationHandler.END

    confirmation = (
        "✅ *Appointment Confirmed!*\n\n"
        f"👤 *Name:* {context.user_data['name']}\n"
        f"🔢 *Age:* {context.user_data['age']}\n"
        f"📋 *Reason:* {context.user_data['reason']}\n"
        f"📱 *Mobile:* {context.user_data['mobile']}\n"
        f"📅 *Date:* {desired_date.strftime('%d/%m/%Y')}\n"
        f"⏰ *Time Slot:* {available_slot}\n\n"
        "Thank you for choosing *Jivandeep Clinic!* 🏥"
    )
    await update.message.reply_text(
        confirmation, parse_mode="Markdown", reply_markup=ReplyKeyboardRemove()
    )

    try:
        await context.bot.send_message(
            chat_id=OWNER_ID,
            text=(
                f"🔔 *New Appointment Booked*\n\n"
                f"👤 {context.user_data['name']} | Age {context.user_data['age']}\n"
                f"📱 {context.user_data['mobile']}\n"
                f"📋 {context.user_data['reason']}\n"
                f"📅 {desired_date.strftime('%d/%m/%Y')} @ {available_slot}"
            ),
            parse_mode="Markdown",
        )
    except Exception:
        pass

    context.user_data.clear()
    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    await update.message.reply_text(
        "❌ Booking cancelled. Send /start to book again.",
        reply_markup=ReplyKeyboardRemove(),
    )
    return ConversationHandler.END


# ══════════════════════════════════════════════
#  OWNER COMMANDS
# ══════════════════════════════════════════════

async def send_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != OWNER_ID:
        return  # Silent ignore
    if os.path.exists(EXCEL_FILE):
        await update.message.reply_document(
            document=open(EXCEL_FILE, "rb"),
            filename=EXCEL_FILE,
            caption="📊 All current appointments.",
        )
    else:
        await update.message.reply_text("ℹ️ No appointments yet.")


async def stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != OWNER_ID:
        return  # Silent ignore
    initialize_excel()
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    total       = ws.max_row - 1
    today_str   = datetime.now().strftime("%d/%m/%Y")
    today_count = sum(1 for row in ws.iter_rows(min_row=2, values_only=True) if row[4] == today_str)
    await update.message.reply_text(
        f"📊 *Appointment Stats*\n\nTotal bookings : {total}\nToday ({today_str}) : {today_count}",
        parse_mode="Markdown",
    )


# ══════════════════════════════════════════════
#  MAIN — uses run_polling() which is Railway compatible
# ══════════════════════════════════════════════

def main():
    if not BOT_TOKEN:
        raise ValueError("BOT_TOKEN environment variable is not set!")
    if not OWNER_ID:
        raise ValueError("OWNER_ID environment variable is not set!")

    initialize_excel()

    app = Application.builder().token(BOT_TOKEN).build()

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            WAITING_FOR_YES : [MessageHandler(filters.TEXT & ~filters.COMMAND, waiting_for_yes)],
            ASK_NAME        : [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_name)],
            ASK_AGE         : [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_age)],
            ASK_REASON      : [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_reason)],
            ASK_MOBILE      : [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_mobile)],
            ASK_DATE        : [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_date)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        allow_reentry=True,
    )

    app.add_handler(conv_handler)
    app.add_handler(CommandHandler("send",  send_excel))
    app.add_handler(CommandHandler("share", send_excel))
    app.add_handler(CommandHandler("stats", stats))

    print("🏥 Jivandeep Clinic Bot is starting...")
    # run_polling() handles its own event loop — works on Railway Python 3.13
    app.run_polling(allowed_updates=Update.ALL_TYPES, drop_pending_updates=True)


if __name__ == "__main__":
    main()
