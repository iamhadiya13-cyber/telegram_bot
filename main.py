"""
Jivandeep Clinic - Enhanced Telegram Appointment Bot
=====================================================
Features:
  ✅ /today /tomorrow owner schedule view
  ✅ View & cancel own appointment (/my_appointment, /cancel_appointment)
  ✅ Appointment reminders (1 day + 1 hour before)
  ✅ Daily summary to owner at 10 PM
  ✅ Gujarati language support
  ✅ Strict age / phone / date validation
Railway + Python 3.13 compatible
"""

import asyncio
import logging
import os
import threading
import json
from datetime import datetime, timedelta, time, date as date_type

import gspread
from google.oauth2.service_account import Credentials
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from telegram import (
    Update,
    ReplyKeyboardMarkup,
    ReplyKeyboardRemove,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
)
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    ConversationHandler,
    CallbackQueryHandler,
    filters,
    ContextTypes,
)

# ─────────────────────────────────────────────
BOT_TOKEN   = os.environ.get("BOT_TOKEN", "")
OWNER_ID    = int(os.environ.get("OWNER_ID", "0"))
SHEET_ID    = os.environ.get("SHEET_ID", "1TPaWWGJS9FwxSPagMRe1e8fPPJgALtM0qWLdDKsbI3I")
SLOTS_FILE  = "slot_config.json"
CREDS_FILE  = "clinicsheet-f489e8632a5f.json"

# Global lock — prevents two users booking the same slot simultaneously
booking_lock = threading.Lock()
# ─────────────────────────────────────────────

# ══════════════════════════════════════════════
#  GOOGLE SHEETS CONNECTION
# ══════════════════════════════════════════════

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

def get_sheet():
    """
    Returns the Google Sheet worksheet.
    Priority:
      1. GOOGLE_CREDS_JSON env variable (Railway) — JSON string
      2. Local file CREDS_FILE (local dev only, never commit to GitHub)
    """
    creds_json_str = os.environ.get("GOOGLE_CREDS_JSON", "").strip()

    if creds_json_str:
        # Load from env variable — json.loads preserves \n in private key exactly
        info = json.loads(creds_json_str)
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    elif os.path.exists(CREDS_FILE):
        creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    else:
        raise RuntimeError(
            "No Google credentials found! "
            "Set GOOGLE_CREDS_JSON environment variable on Railway."
        )

    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SHEET_ID)
    try:
        ws = sh.worksheet("Appointments")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet("Appointments", rows=1000, cols=9)
        ws.append_row(["User ID", "Name", "Age", "Reason", "Mobile",
                       "Date", "Slot Time", "Booking Timestamp", "Status"])
    return ws


def get_all_rows_safe() -> list:
    """Returns rows or empty list on any Google Sheets error."""
    try:
        return get_all_rows()
    except Exception as e:
        logger.error(f"Google Sheets read error: {e}")
        return []


def save_appointment_safe(user_id: int, data: dict) -> bool:
    """Returns True on success, False on failure."""
    try:
        save_appointment(user_id, data)
        return True
    except Exception as e:
        logger.error(f"Google Sheets write error: {e}")
        return False

# ══════════════════════════════════════════════
#  SLOT CONFIG — owner can customize via /setslots
# ══════════════════════════════════════════════

DEFAULT_SLOT_CONFIG = {
    "slot_duration_mins": 30,
    "work_start": "09:00",
    "work_end": "22:00",
    "sunday_end": "13:00",
    "lunch_start": "13:00",
    "lunch_end": "16:00",
    "closed_days": []          # e.g. ["Monday"] to close Mondays
}

def load_slot_config() -> dict:
    if os.path.exists(SLOTS_FILE):
        try:
            with open(SLOTS_FILE) as f:
                cfg = json.load(f)
            # Fill in any missing keys with defaults
            for k, v in DEFAULT_SLOT_CONFIG.items():
                cfg.setdefault(k, v)
            return cfg
        except Exception:
            pass
    return DEFAULT_SLOT_CONFIG.copy()

def save_slot_config(cfg: dict):
    with open(SLOTS_FILE, "w") as f:
        json.dump(cfg, f, indent=2)

# Conversation states
(CHOOSE_LANG, WAITING_FOR_YES, ASK_NAME, ASK_AGE,
 ASK_REASON, ASK_MOBILE, ASK_DATE, ASK_SLOT) = range(8)

# Language strings
STRINGS = {
    "en": {
        "welcome":        "🏥 *Welcome to Jivandeep Clinic!*\n\nWould you like to book an appointment?\nReply *YES* to proceed.",
        "choose_lang":    "🌐 Please choose your language:\n\nભાષા પસંદ કરો:",
        "ask_name":       "👤 Please enter your *Full Name*:",
        "ask_age":        "🔢 Please enter your *Age*:",
        "ask_reason":     "📋 Please enter your *Reason for visit*:",
        "ask_mobile":     "📱 Please enter your *Mobile Number*:\n(10 digits, Indian number e.g. 9876543210)",
        "ask_date":       "📅 Book for *TODAY* or a *SPECIFIC DATE*?\nReply *Today* or enter date as *DD/MM/YYYY*",
        "confirmed":      "✅ *Appointment Confirmed!*\n\n👤 *Name:* {name}\n🔢 *Age:* {age}\n📋 *Reason:* {reason}\n📱 *Mobile:* {mobile}\n📅 *Date:* {date}\n⏰ *Slot:* {slot}\n\nThank you for choosing *Jivandeep Clinic!* 🏥",
        "no_slots":       "😔 No available slots for that date. Please choose another date (DD/MM/YYYY):",
        "past_date":      "⚠️ Cannot book for a past date. Please choose today or a future date:",
        "invalid_date":   "⚠️ Invalid format. Reply *Today* or enter as *DD/MM/YYYY* (e.g. 25/12/2026):",
        "invalid_age":    "⚠️ Invalid age. Please enter a number between 1 and 120:",
        "invalid_mobile": "⚠️ Invalid mobile number. Please enter exactly 10 digits (e.g. 9876543210):",
        "invalid_name":   "⚠️ Name too short. Please enter your full name (at least 2 characters):",
        "closed_sunday":  "⚠️ Clinic is closed on Sundays after 1 PM. Please choose another date:",
        "reply_yes":      "Please reply *YES* to proceed. 😊",
        "cancelled":      "❌ Booking cancelled. Send /start to book again.",
        "no_booking":     "ℹ️ You have no active appointment.",
        "your_appt":      "📋 *Your Appointment*\n\n👤 {name}\n📅 {date}\n⏰ {slot}\n📋 {reason}",
        "cancel_confirm": "Are you sure you want to cancel your appointment on {date} at {slot}?",
        "cancel_done":    "✅ Your appointment has been cancelled successfully.",
        "reminder_1day":  "⏰ *Appointment Reminder*\n\nHello {name}! Your appointment at Jivandeep Clinic is *tomorrow* ({date}) at *{slot}*.\n\nPlease arrive 10 minutes early. 🏥",
        "reminder_1hr":   "⏰ *Appointment Reminder*\n\nHello {name}! Your appointment at Jivandeep Clinic is in *1 hour* at *{slot}*.\n\nPlease get ready! 🏥",
    },
    "gu": {
        "welcome":        "🏥 *જીવનદીપ ક્લિનિકમાં આપનું સ્વાગત છે!*\n\nએપોઇન્ટમેન્ટ બૂક કરવા માટે *YES* જવાબ આપો.",
        "choose_lang":    "🌐 Please choose your language:\n\nભાષા પસંદ કરો:",
        "ask_name":       "👤 કૃપા કરીને તમારું *પૂરું નામ* દાખલ કરો:",
        "ask_age":        "🔢 તમારી *ઉંમર* દાખલ કરો:",
        "ask_reason":     "📋 મુલાકાતનું *કારણ* દાખલ કરો:",
        "ask_mobile":     "📱 તમારો *મોબાઇલ નંબર* દાખલ કરો:\n(10 અંક, દા.ત. 9876543210)",
        "ask_date":       "📅 *આજ* અથવા *ચોક્કસ તારીખ* માટે બૂક કરો?\n*Today* અથવા *DD/MM/YYYY* ફોર્મેટમાં તારીખ લખો",
        "confirmed":      "✅ *એપોઇન્ટમેન્ટ કન્ફર્મ!*\n\n👤 *નામ:* {name}\n🔢 *ઉંમર:* {age}\n📋 *કારણ:* {reason}\n📱 *મોબાઇલ:* {mobile}\n📅 *તારીખ:* {date}\n⏰ *સ્લોટ:* {slot}\n\n*જીવનદીપ ક્લિનિક* પસંદ કરવા બદલ આભાર! 🏥",
        "no_slots":       "😔 આ તારીખ માટે કોઈ સ્લોટ ઉપલબ્ધ નથી. અન્ય તારીખ પસંદ કરો:",
        "past_date":      "⚠️ ભૂતકાળની તારીખ માટે બૂક કરી શકાતું નથી. આજ અથવા ભવિષ્યની તારીખ પસંદ કરો:",
        "invalid_date":   "⚠️ અમાન્ય ફોર્મેટ. *Today* અથવા *DD/MM/YYYY* ફોર્મેટ વાપરો (દા.ત. 25/12/2026):",
        "invalid_age":    "⚠️ અમાન્ય ઉંમર. 1 થી 120 વચ્ચેનો નંબર દાખલ કરો:",
        "invalid_mobile": "⚠️ અમાન્ય મોબાઇલ નંબર. બરાબર 10 અંક દાખલ કરો (દા.ત. 9876543210):",
        "invalid_name":   "⚠️ નામ ખૂબ ટૂંકું છે. ઓછામાં ઓછા 2 અક્ષરોનું પૂરું નામ દાખલ કરો:",
        "closed_sunday":  "⚠️ રવિવારે બપોરે 1 વાગ્યા પછી ક્લિનિક બંધ છે. અન્ય તારીખ પસંદ કરો:",
        "reply_yes":      "એપોઇન્ટમેન્ટ માટે *YES* જવાબ આપો. 😊",
        "cancelled":      "❌ બૂકિંગ રદ. ફરી બૂક કરવા /start મોકલો.",
        "no_booking":     "ℹ️ તમારી કોઈ સક્રિય એપોઇન્ટમેન્ট નથી.",
        "your_appt":      "📋 *તમારી એપોઇન્ટમેન્ટ*\n\n👤 {name}\n📅 {date}\n⏰ {slot}\n📋 {reason}",
        "cancel_confirm": "શું તમે ખરેખર {date} ના {slot} ની એપોઇન્ટમેન્ટ રદ કરવા માંગો છો?",
        "cancel_done":    "✅ તમારી એપોઇન્ટમેન્ટ સફળતાપૂર્વક રદ કરવામાં આવી.",
        "reminder_1day":  "⏰ *એપોઇન્ટમેન્ટ રિમાઇન્ડર*\n\nહેલો {name}! જીવનદીપ ક્લિનિકમાં *આવતીકાલ* ({date}) *{slot}* વાગ્યે એપોઇન્ટમેન્ટ છે.\n\n10 મિનિટ વહેલા આવો. 🏥",
        "reminder_1hr":   "⏰ *એપોઇન્ટમેન્ટ રિમાઇન્ડર*\n\nહેલો {name}! *1 કલાક* પછી *{slot}* વાગ્યે એપોઇન્ટમેન્ટ છે.\n\nતૈયાર થઈ જાઓ! 🏥",
    }
}

def S(user_id, key, **kwargs):
    """Get string for user's language."""
    lang = user_languages.get(user_id, "en")
    text = STRINGS[lang].get(key, STRINGS["en"][key])
    return text.format(**kwargs) if kwargs else text

user_languages = {}  # user_id -> "en" or "gu"

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)


# ══════════════════════════════════════════════
#  GOOGLE SHEETS HELPERS
# ══════════════════════════════════════════════
# Row layout: [User ID, Name, Age, Reason, Mobile, Date, Slot Time, Timestamp, Status]
# Index:          0       1     2    3       4       5      6          7          8

def save_appointment(user_id: int, data: dict):
    """Appends a new appointment row to Google Sheets."""
    ws = get_sheet()
    ws.append_row([
        str(user_id),
        data["name"],
        str(data["age"]),
        data["reason"],
        data["mobile"],
        data["date"].strftime("%d/%m/%Y"),
        data["slot_time"],
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "ACTIVE",
    ])


def get_all_rows() -> list:
    """Returns all data rows (excluding header) as list of lists."""
    ws = get_sheet()
    return ws.get_all_values()[1:]  # skip header row


def get_booked_slots(desired_date: date_type) -> list:
    """Returns list of booked slot strings for a given date."""
    booked = []
    for row in get_all_rows_safe():
        try:
            if len(row) < 9 or row[8] == "CANCELLED":
                continue
            if datetime.strptime(row[5], "%d/%m/%Y").date() == desired_date:
                booked.append(row[6])
        except (ValueError, IndexError):
            continue
    return booked


def get_user_appointment(user_id: int):
    """Returns the latest ACTIVE future appointment for a user, or None."""
    result = None
    for row in get_all_rows_safe():
        try:
            if len(row) < 9:
                continue
            if str(row[0]) == str(user_id) and row[8] == "ACTIVE":
                appt_date = datetime.strptime(row[5], "%d/%m/%Y").date()
                if appt_date >= datetime.now().date():
                    result = row
        except (ValueError, IndexError):
            continue
    return result


def cancel_user_appointment(user_id: int) -> bool:
    """Marks user's latest active future appointment as CANCELLED in Sheets."""
    ws = get_sheet()
    rows = ws.get_all_values()
    cancelled = False
    for i, row in enumerate(rows[1:], start=2):  # start=2 for sheet row number
        try:
            if len(row) < 9:
                continue
            if str(row[0]) == str(user_id) and row[8] == "ACTIVE":
                appt_date = datetime.strptime(row[5], "%d/%m/%Y").date()
                if appt_date >= datetime.now().date():
                    # Column 9 = Status (1-indexed in Sheets)
                    ws.update_cell(i, 9, "CANCELLED")
                    cancelled = True
        except (ValueError, IndexError):
            continue
    return cancelled


def get_appointments_for_date(target_date: date_type) -> list:
    """Returns all active appointments for a given date."""
    appts = []
    for row in get_all_rows_safe():
        try:
            if len(row) < 9 or row[8] == "CANCELLED":
                continue
            if datetime.strptime(row[5], "%d/%m/%Y").date() == target_date:
                appts.append(row)
        except (ValueError, IndexError):
            continue
    return appts


def get_all_future_active_appointments() -> list:
    """Returns all future active appointments for reminders."""
    appts = []
    today = datetime.now().date()
    for row in get_all_rows_safe():
        try:
            if len(row) < 9 or row[8] == "CANCELLED":
                continue
            if datetime.strptime(row[5], "%d/%m/%Y").date() >= today:
                appts.append(row)
        except (ValueError, IndexError):
            continue
    return appts


# ══════════════════════════════════════════════
#  SLOT LOGIC
# ══════════════════════════════════════════════

def get_all_available_slots(desired_date: date_type) -> list:
    """Returns ALL available slots for a date as a list of strings."""
    cfg = load_slot_config()
    def t(s): return datetime.strptime(s, "%H:%M").time()
    WORK_START  = t(cfg["work_start"])
    WORK_END    = t(cfg["work_end"])
    SUNDAY_END  = t(cfg["sunday_end"])
    LUNCH_START = t(cfg["lunch_start"])
    LUNCH_END   = t(cfg["lunch_end"])
    SLOT_MINS   = timedelta(minutes=cfg["slot_duration_mins"])
    is_sunday     = desired_date.weekday() == 6
    effective_end = SUNDAY_END if is_sunday else WORK_END
    booked_slots  = get_booked_slots(desired_date)
    slot_start    = datetime.combine(desired_date, WORK_START)
    end_dt        = datetime.combine(desired_date, effective_end)
    available = []
    while slot_start < end_dt:
        slot_end = slot_start + SLOT_MINS
        if slot_end > end_dt:
            break
        if not is_sunday and LUNCH_START <= slot_start.time() < LUNCH_END:
            slot_start = datetime.combine(desired_date, LUNCH_END)
            continue
        slot_str = f"{slot_start.strftime('%I:%M %p')} - {slot_end.strftime('%I:%M %p')}"
        if slot_str not in booked_slots:
            available.append(slot_str)
        slot_start += SLOT_MINS
    return available


def get_available_slot(desired_date: date_type):
    """Config-driven slot finder. Uses slot_config.json if present."""
    cfg = load_slot_config()

    def t(s): return datetime.strptime(s, "%H:%M").time()

    WORK_START  = t(cfg["work_start"])
    WORK_END    = t(cfg["work_end"])
    SUNDAY_END  = t(cfg["sunday_end"])
    LUNCH_START = t(cfg["lunch_start"])
    LUNCH_END   = t(cfg["lunch_end"])
    SLOT_MINS   = timedelta(minutes=cfg["slot_duration_mins"])

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
    cfg     = load_slot_config()
    now     = datetime.now()
    day_name = desired_date.strftime("%A")  # e.g. "Monday"

    # Check owner-blocked days
    if day_name in cfg.get("closed_days", []):
        return True

    # Sunday half-day check
    if desired_date.weekday() == 6:
        sunday_end = datetime.strptime(cfg["sunday_end"], "%H:%M").time()
        if desired_date == now.date() and now.time() >= sunday_end:
            return True
    return False


def atomic_book_slot(user_id: int, data: dict) -> bool:
    """
    Thread-safe booking — checks slot is still free and saves atomically.
    Returns True if booked successfully, False if slot was taken by someone else.
    """
    with booking_lock:
        # Re-check inside the lock to prevent race condition
        booked = get_booked_slots(data["date"])
        if data["slot_time"] in booked:
            return False  # Slot was just taken by another user
        success = save_appointment_safe(user_id, data)
        return success


# ══════════════════════════════════════════════
#  VALIDATION HELPERS
# ══════════════════════════════════════════════

def validate_age(text: str):
    """Returns (age_int, error_msg). error_msg is None if valid."""
    text = text.strip()
    if not text.isdigit():
        return None, "not_digit"
    age = int(text)
    if age < 1 or age > 120:
        return None, "out_of_range"
    return age, None


def validate_mobile(text: str):
    """
    Valid Indian mobile: exactly 10 digits, starts with 6-9.
    Also accepts +91 prefix.
    Returns (cleaned_number, error_msg).
    """
    text = text.strip()
    # Remove +91 or 91 prefix
    if text.startswith("+91"):
        text = text[3:]
    elif text.startswith("91") and len(text) == 12:
        text = text[2:]
    text = text.replace(" ", "").replace("-", "")
    if not text.isdigit():
        return None, "not_digit"
    if len(text) != 10:
        return None, "not_10_digits"
    if text[0] not in "6789":
        return None, "invalid_start"
    return text, None


def validate_date(text: str):
    """Returns (date_obj, error_msg)."""
    text = text.strip()
    if text.lower() == "today":
        return datetime.now().date(), None
    try:
        d = datetime.strptime(text, "%d/%m/%Y").date()
        return d, None
    except ValueError:
        return None, "invalid_format"


# ══════════════════════════════════════════════
#  LANGUAGE SELECTION
# ══════════════════════════════════════════════

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    keyboard = [[
        InlineKeyboardButton("🇬🇧 English", callback_data="lang_en"),
        InlineKeyboardButton("🇮🇳 ગુજરાતી", callback_data="lang_gu"),
    ]]
    await update.message.reply_text(
        STRINGS["en"]["choose_lang"],
        reply_markup=InlineKeyboardMarkup(keyboard),
    )
    return CHOOSE_LANG


async def choose_language(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()
    lang = "en" if query.data == "lang_en" else "gu"
    user_id = query.from_user.id
    user_languages[user_id] = lang

    keyboard = [["YES"]]
    await query.edit_message_text(
        S(user_id, "welcome"),
        parse_mode="Markdown",
    )
    await context.bot.send_message(
        chat_id=query.message.chat_id,
        text=S(user_id, "welcome"),
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True),
    )
    return WAITING_FOR_YES


# ══════════════════════════════════════════════
#  CONVERSATION HANDLERS
# ══════════════════════════════════════════════

async def waiting_for_yes(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    user_id = update.effective_user.id
    if update.message.text.strip().upper() == "YES":
        await update.message.reply_text(
            S(user_id, "ask_name"),
            parse_mode="Markdown",
            reply_markup=ReplyKeyboardRemove(),
        )
        return ASK_NAME
    keyboard = [["YES"]]
    await update.message.reply_text(
        S(user_id, "reply_yes"),
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True),
    )
    return WAITING_FOR_YES


async def ask_name(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    user_id = update.effective_user.id
    name = update.message.text.strip()
    # Validation: at least 2 chars
    # isalpha() works for ALL unicode including Gujarati (ગુજરાતી), Hindi, etc.
    # Also allow spaces and dots (for names like "Dr. Shah")
    allowed = all(c.isalpha() or c.isspace() or c == '.' for c in name)
    if len(name) < 2 or not allowed:
        await update.message.reply_text(S(user_id, "invalid_name"))
        return ASK_NAME
    context.user_data["name"] = name
    await update.message.reply_text(S(user_id, "ask_age"), parse_mode="Markdown")
    return ASK_AGE


async def ask_age(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    user_id = update.effective_user.id
    age, err = validate_age(update.message.text)
    if err:
        await update.message.reply_text(S(user_id, "invalid_age"))
        return ASK_AGE
    context.user_data["age"] = age
    await update.message.reply_text(S(user_id, "ask_reason"), parse_mode="Markdown")
    return ASK_REASON


async def ask_reason(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    user_id = update.effective_user.id
    reason = update.message.text.strip()
    if len(reason) < 3:
        await update.message.reply_text("⚠️ Please describe your reason (at least 3 characters):")
        return ASK_REASON
    context.user_data["reason"] = reason
    await update.message.reply_text(S(user_id, "ask_mobile"), parse_mode="Markdown")
    return ASK_MOBILE


async def ask_mobile(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    user_id = update.effective_user.id
    mobile, err = validate_mobile(update.message.text)
    if err:
        await update.message.reply_text(S(user_id, "invalid_mobile"))
        return ASK_MOBILE
    context.user_data["mobile"] = mobile
    keyboard = [["Today"]]
    await update.message.reply_text(
        S(user_id, "ask_date"),
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True),
    )
    return ASK_DATE


async def ask_date(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    user_id = update.effective_user.id
    text  = update.message.text.strip()
    today = datetime.now().date()

    desired_date, err = validate_date(text)
    if err:
        await update.message.reply_text(
            S(user_id, "invalid_date"), parse_mode="Markdown"
        )
        return ASK_DATE

    if desired_date < today:
        await update.message.reply_text(
            S(user_id, "past_date"), reply_markup=ReplyKeyboardRemove()
        )
        return ASK_DATE

    if is_clinic_closed(desired_date):
        await update.message.reply_text(
            S(user_id, "closed_sunday"), reply_markup=ReplyKeyboardRemove()
        )
        return ASK_DATE

    # Get ALL available slots and show picker
    try:
        available_slots = get_all_available_slots(desired_date)
    except Exception as e:
        logger.error(f"Sheets error in ask_date: {e}")
        await update.message.reply_text(
            "⚠️ Could not connect to booking system. Please try again in a moment.",
            reply_markup=ReplyKeyboardRemove(),
        )
        return ConversationHandler.END
    available_slots = available_slots  # already assigned above
    if not available_slots:
        tomorrow = desired_date + timedelta(days=1)
        tomorrow_slots = get_all_available_slots(tomorrow)
        if tomorrow_slots:
            keyboard = [[tomorrow.strftime("%d/%m/%Y"), "Other Date"]]
            d = desired_date.strftime("%d/%m/%Y")
            t_str = tomorrow.strftime("%d/%m/%Y")
            await update.message.reply_text(
                "No slots available for " + d + ".\n\n"
                "Tomorrow (" + t_str + ") has slots!\n"
                "Tap tomorrow or enter another date (DD/MM/YYYY):",
                reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True),
            )
        else:
            await update.message.reply_text(
                S(user_id, "no_slots"), reply_markup=ReplyKeyboardRemove()
            )
        return ASK_DATE

    # Save date, show slot buttons
    context.user_data["date"] = desired_date
    await show_slot_picker(update, context, available_slots, desired_date)
    return ASK_SLOT

    await update.message.reply_text(
        S(user_id, "confirmed",
          name=context.user_data["name"],
          age=context.user_data["age"],
          reason=context.user_data["reason"],
          mobile=context.user_data["mobile"],
          date=desired_date.strftime("%d/%m/%Y"),
          slot=""),
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardRemove(),
    )

    context.user_data.clear()
    return ConversationHandler.END


async def show_slot_picker(update, context, slots: list, desired_date):
    """Shows inline keyboard with available time slots grouped by period."""
    user_id = update.effective_user.id
    # Group: Morning (before 12), Afternoon (12-17), Evening (17+)
    morning, afternoon, evening = [], [], []
    for slot in slots:
        start_str = slot.split(" - ")[0].strip()
        h = datetime.strptime(start_str, "%I:%M %p").hour
        if h < 12:
            morning.append(slot)
        elif h < 17:
            afternoon.append(slot)
        else:
            evening.append(slot)

    keyboard = []
    if morning:
        keyboard.append([InlineKeyboardButton("🌅 Morning", callback_data="grp_morning")])
        # 2 slots per row
        for i in range(0, len(morning), 2):
            row = [InlineKeyboardButton(s, callback_data="slot_" + s) for s in morning[i:i+2]]
            keyboard.append(row)
    if afternoon:
        keyboard.append([InlineKeyboardButton("☀️ Afternoon", callback_data="grp_afternoon")])
        for i in range(0, len(afternoon), 2):
            row = [InlineKeyboardButton(s, callback_data="slot_" + s) for s in afternoon[i:i+2]]
            keyboard.append(row)
    if evening:
        keyboard.append([InlineKeyboardButton("🌆 Evening", callback_data="grp_evening")])
        for i in range(0, len(evening), 2):
            row = [InlineKeyboardButton(s, callback_data="slot_" + s) for s in evening[i:i+2]]
            keyboard.append(row)

    date_str = desired_date.strftime("%d/%m/%Y")
    total    = len(slots)
    await update.message.reply_text(
        S(user_id, "pick_slot", date=date_str) + f"\n\n_{total} slots available_",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )


async def ask_slot(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Handles slot selection via inline keyboard button."""
    query   = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    data    = query.data

    # Ignore group header button taps
    if data.startswith("grp_"):
        return ASK_SLOT

    if not data.startswith("slot_"):
        await query.edit_message_text(S(user_id, "invalid_slot"))
        return ASK_SLOT

    chosen_slot = data[5:]  # strip "slot_" prefix
    desired_date = context.user_data.get("date")

    if not desired_date:
        await query.edit_message_text("Session expired. Please /start again.")
        return ConversationHandler.END

    # Verify slot still available
    still_available = get_all_available_slots(desired_date)
    if chosen_slot not in still_available:
        # Slot was just taken — refresh picker
        if still_available:
            user_id = query.from_user.id
            morning, afternoon, evening = [], [], []
            for slot in still_available:
                h = datetime.strptime(slot.split(" - ")[0].strip(), "%I:%M %p").hour
                if h < 12: morning.append(slot)
                elif h < 17: afternoon.append(slot)
                else: evening.append(slot)
            keyboard = []
            for label, group in [("🌅 Morning","grp_m"),("☀️ Afternoon","grp_a"),("🌆 Evening","grp_e")]:
                grp = morning if "M" in label else (afternoon if "A" in label else evening)
                if grp:
                    keyboard.append([InlineKeyboardButton(label, callback_data=label)])
                    for i in range(0, len(grp), 2):
                        keyboard.append([InlineKeyboardButton(s, callback_data="slot_"+s) for s in grp[i:i+2]])
            await query.edit_message_text(
                S(user_id, "slot_taken"),
                reply_markup=InlineKeyboardMarkup(keyboard),
            )
        else:
            tomorrow = desired_date + timedelta(days=1)
            t_slots  = get_all_available_slots(tomorrow)
            if t_slots:
                await query.edit_message_text(
                    "⚠️ All slots just filled for this date!\n"
                    "Tomorrow (" + tomorrow.strftime("%d/%m/%Y") + ") has slots.\n"
                    "Please type " + tomorrow.strftime("%d/%m/%Y") + " to book tomorrow.",
                )
                context.user_data["date"] = None
                return ASK_DATE
            else:
                await query.edit_message_text(
                    "😔 All slots are full. Please /start and choose another date."
                )
                return ConversationHandler.END
        return ASK_SLOT

    # Lock and book
    context.user_data["slot_time"] = chosen_slot
    try:
        booked = atomic_book_slot(user_id, context.user_data)
    except PermissionError:
        await query.edit_message_text("⚠️ Could not save. Please try again.")
        return ConversationHandler.END

    if not booked:
        await query.edit_message_text(S(user_id, "slot_taken"))
        return ASK_SLOT

    # Confirm
    confirmation = (
        S(user_id, "confirmed",
          name=context.user_data["name"],
          age=context.user_data["age"],
          reason=context.user_data["reason"],
          mobile=context.user_data["mobile"],
          date=desired_date.strftime("%d/%m/%Y"),
          slot=chosen_slot)
    )
    await query.edit_message_text(confirmation, parse_mode="Markdown")

    # Notify owner
    try:
        await context.bot.send_message(
            chat_id=OWNER_ID,
            text=(
                f"🔔 *New Appointment*\n\n"
                f"👤 {context.user_data['name']} | Age {context.user_data['age']}\n"
                f"📱 {context.user_data['mobile']}\n"
                f"📋 {context.user_data['reason']}\n"
                f"📅 {desired_date.strftime('%d/%m/%Y')} @ {chosen_slot}"
            ),
            parse_mode="Markdown",
        )
    except Exception:
        pass

    context.user_data.clear()
    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    user_id = update.effective_user.id
    context.user_data.clear()
    await update.message.reply_text(
        S(user_id, "cancelled"), reply_markup=ReplyKeyboardRemove()
    )
    return ConversationHandler.END


# ══════════════════════════════════════════════
#  PATIENT SELF-SERVICE
# ══════════════════════════════════════════════

async def my_appointment(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    appt = get_user_appointment(user_id)
    if not appt:
        await update.message.reply_text(S(user_id, "no_booking"))
        return
    await update.message.reply_text(
        S(user_id, "your_appt",
          name=appt[1], date=appt[5], slot=appt[6], reason=appt[3]),
        parse_mode="Markdown",
    )


async def cancel_appointment_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    appt = get_user_appointment(user_id)
    if not appt:
        await update.message.reply_text(S(user_id, "no_booking"))
        return
    keyboard = [[
        InlineKeyboardButton("✅ Yes, Cancel", callback_data="confirm_cancel"),
        InlineKeyboardButton("❌ No, Keep it", callback_data="keep_appt"),
    ]]
    await update.message.reply_text(
        S(user_id, "cancel_confirm", date=appt[5], slot=appt[6]),
        reply_markup=InlineKeyboardMarkup(keyboard),
    )


async def handle_cancel_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id

    if query.data == "confirm_cancel":
        success = cancel_user_appointment(user_id)
        if success:
            await query.edit_message_text(S(user_id, "cancel_done"))
            try:
                appt = get_user_appointment(user_id)
                name = appt[1] if appt else "Patient"
                await context.bot.send_message(
                    chat_id=OWNER_ID,
                    text=f"🚫 *Appointment Cancelled*\n\n👤 Patient cancelled their appointment.",
                    parse_mode="Markdown",
                )
            except Exception:
                pass
        else:
            await query.edit_message_text("⚠️ Could not cancel. Please try again.")
    else:
        await query.edit_message_text("✅ Your appointment is kept. See you soon! 🏥")


# ══════════════════════════════════════════════
#  OWNER COMMANDS
# ══════════════════════════════════════════════

def format_schedule(appts: list, title: str) -> str:
    if not appts:
        return f"📅 *{title}*\n\nNo appointments."
    # Sort by slot time
    appts_sorted = sorted(appts, key=lambda r: r[6])
    lines = [f"📅 *{title}* — {len(appts_sorted)} appointment(s)\n"]
    for i, row in enumerate(appts_sorted, 1):
        lines.append(f"{i}. ⏰ {row[6]}\n   👤 {row[1]} | Age {row[2]}\n   📱 {row[4]} | 📋 {row[3]}\n")
    return "\n".join(lines)


async def today_schedule(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != OWNER_ID:
        return
    appts = get_appointments_for_date(datetime.now().date())
    msg   = format_schedule(appts, f"Today's Schedule ({datetime.now().strftime('%d/%m/%Y')})")
    await update.message.reply_text(msg, parse_mode="Markdown")


async def tomorrow_schedule(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != OWNER_ID:
        return
    tomorrow = datetime.now().date() + timedelta(days=1)
    appts    = get_appointments_for_date(tomorrow)
    msg      = format_schedule(appts, f"Tomorrow's Schedule ({tomorrow.strftime('%d/%m/%Y')})")
    await update.message.reply_text(msg, parse_mode="Markdown")


async def send_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != OWNER_ID:
        return
    await update.message.reply_text(
        "📊 View all appointments live on Google Sheets:\n\n"
        "https://docs.google.com/spreadsheets/d/" + SHEET_ID + "/edit",
        parse_mode="Markdown",
    )


async def stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != OWNER_ID:
        return
    today_str = datetime.now().strftime("%d/%m/%Y")
    total, today_count, cancelled_count = 0, 0, 0
    for row in get_all_rows_safe():
        if len(row) >= 9 and row[8] == "CANCELLED":
            cancelled_count += 1
            continue
        total += 1
        if len(row) >= 6 and row[5] == today_str:
            today_count += 1
    await update.message.reply_text(
        f"📊 *Appointment Stats*\n\n"
        f"✅ Total active  : {total}\n"
        f"❌ Cancelled     : {cancelled_count}\n"
        f"📅 Today ({today_str}): {today_count}",
        parse_mode="Markdown",
    )


# Owner slot config states
SET_SLOTS_MENU, SET_DURATION, SET_WORK_START, SET_WORK_END, SET_LUNCH, SET_CLOSE_DAY = range(10, 16)

async def show_slot_config(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Owner: /setslots — show current config and options."""
    if update.effective_user.id != OWNER_ID:
        return
    cfg = load_slot_config()
    closed = ", ".join(cfg["closed_days"]) if cfg["closed_days"] else "None"
    summary = (
        f"⚙️ *Current Slot Settings*\n\n"
        f"⏱ Slot duration   : {cfg['slot_duration_mins']} mins\n"
        f"🌅 Work start      : {cfg['work_start']}\n"
        f"🌆 Work end        : {cfg['work_end']}\n"
        f"☀️ Sunday end      : {cfg['sunday_end']}\n"
        f"🍽 Lunch break     : {cfg['lunch_start']} – {cfg['lunch_end']}\n"
        f"🚫 Closed days     : {closed}\n\n"
        f"Use buttons below to change settings:"
    )
    keyboard = [
        [InlineKeyboardButton("⏱ Slot Duration", callback_data="cfg_duration")],
        [InlineKeyboardButton("🌅 Work Start", callback_data="cfg_wstart"),
         InlineKeyboardButton("🌆 Work End",   callback_data="cfg_wend")],
        [InlineKeyboardButton("🍽 Lunch Break", callback_data="cfg_lunch")],
        [InlineKeyboardButton("🚫 Close a Day", callback_data="cfg_closeday"),
         InlineKeyboardButton("✅ Open a Day",  callback_data="cfg_openday")],
        [InlineKeyboardButton("🔄 Reset to Default", callback_data="cfg_reset")],
    ]
    await update.message.reply_text(
        summary, parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )


async def slot_config_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle slot config button presses."""
    query = update.callback_query
    await query.answer()
    if query.from_user.id != OWNER_ID:
        return

    data = query.data
    cfg  = load_slot_config()

    if data == "cfg_reset":
        save_slot_config(DEFAULT_SLOT_CONFIG.copy())
        await query.edit_message_text("✅ Slot config reset to default.")

    elif data == "cfg_duration":
        context.user_data["cfg_editing"] = "duration"
        await query.edit_message_text(
            "⏱ Enter new slot duration in minutes\n"
            "Examples: *15*, *20*, *30*, *45*, *60*",
            parse_mode="Markdown"
        )

    elif data == "cfg_wstart":
        context.user_data["cfg_editing"] = "wstart"
        await query.edit_message_text(
            "🌅 Enter new work start time (24hr format)\nExample: *08:00* or *09:00*",
            parse_mode="Markdown"
        )

    elif data == "cfg_wend":
        context.user_data["cfg_editing"] = "wend"
        await query.edit_message_text(
            "🌆 Enter new work end time (24hr format)\nExample: *21:00* or *22:00*",
            parse_mode="Markdown"
        )

    elif data == "cfg_lunch":
        context.user_data["cfg_editing"] = "lunch"
        await query.edit_message_text(
            "🍽 Enter lunch break as START-END (24hr)\nExample: *13:00-16:00*",
            parse_mode="Markdown"
        )

    elif data == "cfg_closeday":
        context.user_data["cfg_editing"] = "closeday"
        days = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
        keyboard = [[InlineKeyboardButton(d, callback_data=f"close_{d}")] for d in days]
        await query.edit_message_text(
            "🚫 Which day to close?",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

    elif data == "cfg_openday":
        closed = cfg.get("closed_days", [])
        if not closed:
            await query.edit_message_text("No days are currently closed.")
            return
        keyboard = [[InlineKeyboardButton(d, callback_data=f"open_{d}")] for d in closed]
        await query.edit_message_text(
            "✅ Which day to reopen?",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

    elif data.startswith("close_"):
        day = data[6:]
        if day not in cfg["closed_days"]:
            cfg["closed_days"].append(day)
            save_slot_config(cfg)
        await query.edit_message_text(f"🚫 *{day}* is now closed.", parse_mode="Markdown")

    elif data.startswith("open_"):
        day = data[5:]
        if day in cfg["closed_days"]:
            cfg["closed_days"].remove(day)
            save_slot_config(cfg)
        await query.edit_message_text(f"✅ *{day}* is now open.", parse_mode="Markdown")


async def handle_cfg_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle text replies for slot config editing (owner only)."""
    if update.effective_user.id != OWNER_ID:
        return
    editing = context.user_data.get("cfg_editing")
    if not editing:
        return

    text = update.message.text.strip()
    cfg  = load_slot_config()

    if editing == "duration":
        if not text.isdigit() or not (5 <= int(text) <= 120):
            await update.message.reply_text("⚠️ Enter a number between 5 and 120 minutes:")
            return
        cfg["slot_duration_mins"] = int(text)
        save_slot_config(cfg)
        await update.message.reply_text(f"✅ Slot duration set to *{text} minutes*.", parse_mode="Markdown")

    elif editing in ("wstart", "wend"):
        try:
            datetime.strptime(text, "%H:%M")
        except ValueError:
            await update.message.reply_text("⚠️ Use HH:MM format e.g. 09:00")
            return
        key = "work_start" if editing == "wstart" else "work_end"
        cfg[key] = text
        save_slot_config(cfg)
        label = "Work start" if editing == "wstart" else "Work end"
        await update.message.reply_text(f"✅ {label} set to *{text}*.", parse_mode="Markdown")

    elif editing == "lunch":
        parts = text.split("-")
        if len(parts) != 2:
            await update.message.reply_text("⚠️ Use format START-END e.g. 13:00-16:00")
            return
        try:
            datetime.strptime(parts[0].strip(), "%H:%M")
            datetime.strptime(parts[1].strip(), "%H:%M")
        except ValueError:
            await update.message.reply_text("⚠️ Use format START-END e.g. 13:00-16:00")
            return
        cfg["lunch_start"] = parts[0].strip()
        cfg["lunch_end"]   = parts[1].strip()
        save_slot_config(cfg)
        await update.message.reply_text(
            f"✅ Lunch break set to *{parts[0].strip()} – {parts[1].strip()}*.",
            parse_mode="Markdown"
        )

    context.user_data.pop("cfg_editing", None)
    # Show updated config
    closed = ", ".join(cfg["closed_days"]) if cfg["closed_days"] else "None"
    await update.message.reply_text(
        f"⚙️ *Updated Settings*\n\n"
        f"⏱ Duration  : {cfg['slot_duration_mins']} mins\n"
        f"🌅 Start     : {cfg['work_start']}\n"
        f"🌆 End       : {cfg['work_end']}\n"
        f"🍽 Lunch     : {cfg['lunch_start']} – {cfg['lunch_end']}\n"
        f"🚫 Closed    : {closed}",
        parse_mode="Markdown"
    )


async def search_patient(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != OWNER_ID:
        return
    if not context.args:
        await update.message.reply_text("Usage: /search <name>\nExample: /search Rahul")
        return
    query_name = " ".join(context.args).lower()
    results = []
    for row in get_all_rows_safe():
        if len(row) >= 2 and row[1] and query_name in str(row[1]).lower():
            results.append(row)
    if not results:
        await update.message.reply_text(f"No results for '{query_name}'.")
        return
    lines = [f"🔍 *Search: {query_name}* — {len(results)} found\n"]
    for r in results:
        status = "✅" if (len(r) >= 9 and r[8] == "ACTIVE") else "❌"
        lines.append(f"{status} {r[1]} | {r[5]} @ {r[6]}\n   📱 {r[4]} | 📋 {r[3]}")
    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")


# ══════════════════════════════════════════════
#  SCHEDULED JOBS
# ══════════════════════════════════════════════

async def send_daily_summary(bot):
    """Sends daily summary to owner at 10 PM."""
    try:
        today    = datetime.now().date()
        tomorrow = today + timedelta(days=1)
        today_appts    = get_appointments_for_date(today)
        tomorrow_appts = get_appointments_for_date(tomorrow)

        msg = (
            f"🌙 *Daily Summary — {today.strftime('%d/%m/%Y')}*\n\n"
            f"Today's patients    : {len(today_appts)}\n"
            f"Tomorrow's patients : {len(tomorrow_appts)}\n"
        )
        if tomorrow_appts:
            first = sorted(tomorrow_appts, key=lambda r: r[6])[0]
            msg += f"\nFirst slot tomorrow : {first[6]} — {first[1]}"

        await bot.send_message(chat_id=OWNER_ID, text=msg, parse_mode="Markdown")
    except Exception as e:
        logger.error(f"Daily summary error: {e}")


async def send_reminders(bot):
    """Checks every 30 min and sends 1-day and 1-hour reminders."""
    try:
        now       = datetime.now()
        tomorrow  = (now + timedelta(days=1)).date()
        appts     = get_all_future_active_appointments()

        for appt in appts:
            user_id    = int(appt[0]) if str(appt[0]).isdigit() else appt[0]
            name       = appt[1]
            appt_date  = datetime.strptime(appt[5], "%d/%m/%Y").date()
            slot_start = appt[6].split(" - ")[0].strip()  # e.g. "10:30 AM"

            try:
                appt_dt = datetime.combine(
                    appt_date,
                    datetime.strptime(slot_start, "%I:%M %p").time()
                )
            except ValueError:
                continue

            diff_hours = (appt_dt - now).total_seconds() / 3600

            # 1-day reminder: between 23h and 25h before
            if 23 <= diff_hours <= 25:
                lang = user_languages.get(user_id, "en")
                msg  = STRINGS[lang]["reminder_1day"].format(
                    name=name, date=appt[5], slot=appt[6]
                )
                try:
                    await bot.send_message(chat_id=user_id, text=msg, parse_mode="Markdown")
                except Exception:
                    pass

            # 1-hour reminder: between 0.9h and 1.1h before
            elif 0.9 <= diff_hours <= 1.1:
                lang = user_languages.get(user_id, "en")
                msg  = STRINGS[lang]["reminder_1hr"].format(
                    name=name, slot=appt[6]
                )
                try:
                    await bot.send_message(chat_id=user_id, text=msg, parse_mode="Markdown")
                except Exception:
                    pass

    except Exception as e:
        logger.error(f"Reminder error: {e}")


# ══════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════

async def post_init(app):
    """Set bot command menus after bot starts."""
    from telegram import BotCommand, BotCommandScopeDefault, BotCommandScopeChat

    # All users see these commands when they type /
    user_commands = [
        BotCommand("start",               "Book an appointment"),
        BotCommand("my_appointment",      "View my current appointment"),
        BotCommand("cancel_appointment",  "Cancel my appointment"),
        BotCommand("cancel",              "Cancel ongoing booking"),
    ]
    await app.bot.set_my_commands(user_commands, scope=BotCommandScopeDefault())

    # Owner also sees admin commands
    try:
        owner_commands = user_commands + [
            BotCommand("today",    "Today's appointment schedule"),
            BotCommand("tomorrow", "Tomorrow's schedule"),
            BotCommand("stats",    "Booking statistics"),
            BotCommand("send",     "Get appointments Excel file"),
            BotCommand("search",   "Search patient by name"),
            BotCommand("setslots", "Configure slot timings"),
        ]
        await app.bot.set_my_commands(
            owner_commands,
            scope=BotCommandScopeChat(chat_id=OWNER_ID)
        )
    except Exception:
        pass  # Owner hasn't started bot yet


def main():
    if not BOT_TOKEN:
        raise ValueError("BOT_TOKEN environment variable not set!")
    if not OWNER_ID:
        raise ValueError("OWNER_ID environment variable not set!")

    app = Application.builder().token(BOT_TOKEN).post_init(post_init).build()

    async def error_handler(update, context):
        logger.error(f"Exception: {context.error}", exc_info=context.error)
        if update and update.effective_message:
            try:
                await update.effective_message.reply_text(
                    "⚠️ Something went wrong. Please try again or type /start."
                )
            except Exception:
                pass
    app.add_error_handler(error_handler)

    # Conversation
    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            CHOOSE_LANG     : [CallbackQueryHandler(choose_language, pattern="^lang_")],
            WAITING_FOR_YES : [MessageHandler(filters.TEXT & ~filters.COMMAND, waiting_for_yes)],
            ASK_NAME        : [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_name)],
            ASK_AGE         : [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_age)],
            ASK_REASON      : [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_reason)],
            ASK_MOBILE      : [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_mobile)],
            ASK_DATE        : [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_date)],
            ASK_SLOT        : [CallbackQueryHandler(ask_slot, pattern="^(slot_|grp_)")],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        allow_reentry=True,
    )

    app.add_handler(conv_handler)

    # Patient commands
    app.add_handler(CommandHandler("my_appointment",     my_appointment))
    app.add_handler(CommandHandler("cancel_appointment", cancel_appointment_cmd))
    app.add_handler(CallbackQueryHandler(handle_cancel_callback, pattern="^(confirm_cancel|keep_appt)$"))

    # Owner commands
    app.add_handler(CommandHandler("today",    today_schedule))
    app.add_handler(CommandHandler("tomorrow", tomorrow_schedule))
    app.add_handler(CommandHandler("send",     send_excel))
    app.add_handler(CommandHandler("share",    send_excel))
    app.add_handler(CommandHandler("stats",    stats))
    app.add_handler(CommandHandler("search",   search_patient))
    app.add_handler(CommandHandler("setslots", show_slot_config))
    app.add_handler(CallbackQueryHandler(slot_config_callback, pattern="^(cfg_|close_|open_)"))
    app.add_handler(MessageHandler(
        filters.TEXT & ~filters.COMMAND & filters.User(OWNER_ID),
        handle_cfg_text
    ))

    # Scheduler
    scheduler = AsyncIOScheduler()
    # Daily summary at 10 PM
    scheduler.add_job(
        send_daily_summary, "cron", hour=22, minute=0,
        args=[app.bot]
    )
    # Reminders check every 30 minutes
    scheduler.add_job(
        send_reminders, "interval", minutes=30,
        args=[app.bot]
    )
    scheduler.start()

    print("🏥 Jivandeep Clinic Bot is starting...")
    app.run_polling(allowed_updates=Update.ALL_TYPES, drop_pending_updates=True)


if __name__ == "__main__":
    main()
