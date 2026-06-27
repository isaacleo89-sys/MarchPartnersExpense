"""
March Partners Expense Bot
──────────────────────────
Send a receipt photo — the bot guides you through 5 steps via buttons and text:
  1. Entity   (tap a button)
  2. Amount   (type a number)
  3. Payer    (tap a button: IL / JK / Fund)
  4. Date     (tap Today or type a custom date)
  5. Purpose  (type a description)

The receipt is then saved to Expenses/{Entity}/{YYYYMMMDD_Purpose}.jpg in
SharePoint, and a new row is appended to the correct tab in the Excel tracker.
"""

import asyncio
import logging
import re
import unicodedata
from datetime import datetime, date as date_type

from telegram import (
    Update,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
)
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ConversationHandler,
    MessageHandler,
    filters,
    ContextTypes,
)

from config import (
    TELEGRAM_TOKEN,
    AUTHORIZED_USER_IDS,
    ENTITY_MAP,
    ENTITY_FULL_NAMES,
    ENTITY_FOLDER_NAMES,
    EXCEL_FILE_PATH,
    RECEIPTS_BASE_PATH,
    PAYERS,
)
from graph_client import GraphClient
from excel_manager import add_expense

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger(__name__)

# ── Conversation states ───────────────────────────────────────────────────────
(
    SELECTING_ENTITY,
    ENTERING_AMOUNT,
    SELECTING_PAYER,
    ENTERING_DATE,
    ENTERING_PURPOSE,
) = range(5)

# ── Shared Graph client ───────────────────────────────────────────────────────
graph = GraphClient()


# ── Authorisation ─────────────────────────────────────────────────────────────
def _is_authorised(user_id: int) -> bool:
    return not AUTHORIZED_USER_IDS or user_id in AUTHORIZED_USER_IDS


# ── Keyboard builders ─────────────────────────────────────────────────────────

def _entity_keyboard() -> InlineKeyboardMarkup:
    """One button per entity, two columns."""
    sheets = list(dict.fromkeys(ENTITY_MAP.values()))   # unique, ordered
    buttons = [
        InlineKeyboardButton(name, callback_data=f"entity:{name}")
        for name in sheets
    ]
    # Arrange in rows of 2
    rows = [buttons[i:i + 2] for i in range(0, len(buttons), 2)]
    return InlineKeyboardMarkup(rows)


def _payer_keyboard() -> InlineKeyboardMarkup:
    buttons = [
        InlineKeyboardButton(p, callback_data=f"payer:{p}") for p in PAYERS
    ]
    return InlineKeyboardMarkup([buttons])


def _date_keyboard(today: datetime) -> InlineKeyboardMarkup:
    label = f"✓  Today — {today.strftime('%-d %b %Y')}"
    return InlineKeyboardMarkup([
        [InlineKeyboardButton(label, callback_data="date:TODAY")]
    ])


# ── Filename helper ───────────────────────────────────────────────────────────

def _receipt_filename(expense_date: datetime, purpose: str) -> str:
    date_str = expense_date.strftime("%Y%b%d")           # e.g. 2026Jun27
    safe = unicodedata.normalize("NFKD", purpose).encode("ascii", "ignore").decode()
    safe = re.sub(r"[^\w\s-]", "", safe).strip()
    safe = re.sub(r"\s+", "_", safe)[:60]
    return f"{date_str}_{safe}.jpg"


# ── Date parsing ──────────────────────────────────────────────────────────────
_DATE_FORMATS = [
    "%d %b %Y", "%d %B %Y",   # 27 Jun 2026 / 27 June 2026
    "%d %b",    "%d %B",       # 27 Jun / 27 June  (year inferred = current)
    "%d/%m/%Y", "%d-%m-%Y",    # 27/06/2026
    "%Y-%m-%d",                # 2026-06-27
]

def _parse_date(text: str) -> datetime | None:
    text = text.strip()
    for fmt in _DATE_FORMATS:
        try:
            dt = datetime.strptime(text, fmt)
            if dt.year == 1900:               # no year in format
                dt = dt.replace(year=datetime.now().year)
            return dt
        except ValueError:
            continue
    return None


# ── Command handlers ──────────────────────────────────────────────────────────

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not _is_authorised(update.effective_user.id):
        return
    await update.message.reply_text(
        "👋 *March Partners Expense Bot*\n\n"
        "Send me a photo of your receipt and I'll guide you through the rest.\n\n"
        "Commands:\n"
        "• /start — show this message\n"
        "• /cancel — cancel the current entry",
        parse_mode="Markdown",
    )


async def cmd_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    await update.message.reply_text("❌ Entry cancelled.")
    return ConversationHandler.END


# ── Step 1 — Photo received → ask for entity ──────────────────────────────────

async def step_photo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not _is_authorised(update.effective_user.id):
        return ConversationHandler.END

    # Download and store the photo
    photo = update.message.photo[-1]
    photo_file = await photo.get_file()
    image_bytes = bytes(await photo_file.download_as_bytearray())
    context.user_data["image_bytes"] = image_bytes
    context.user_data.pop("entity",  None)
    context.user_data.pop("amount",  None)
    context.user_data.pop("payer",   None)
    context.user_data.pop("expense_date", None)
    context.user_data.pop("purpose", None)

    await update.message.reply_text(
        "📁 *Which entity should this be charged to?*",
        reply_markup=_entity_keyboard(),
        parse_mode="Markdown",
    )
    return SELECTING_ENTITY


# ── Step 2 — Entity selected → ask for amount ─────────────────────────────────

async def step_entity(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()

    sheet_name = query.data.split("entity:", 1)[1]
    context.user_data["entity"] = sheet_name

    await query.edit_message_text(
        f"📁 Entity: *{sheet_name}*\n\n"
        "💰 *Enter the amount (SGD):*",
        parse_mode="Markdown",
    )
    return ENTERING_AMOUNT


# ── Step 3 — Amount entered → ask for payer ───────────────────────────────────

async def step_amount(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    text = update.message.text.strip()
    cleaned = re.sub(r"[^\d.]", "", text)
    try:
        amount = float(cleaned)
        if amount <= 0:
            raise ValueError
    except ValueError:
        await update.message.reply_text(
            "❌ Please enter a valid amount, e.g. `123.45`",
            parse_mode="Markdown",
        )
        return ENTERING_AMOUNT

    context.user_data["amount"] = amount
    entity = context.user_data["entity"]

    await update.message.reply_text(
        f"📁 Entity: *{entity}*\n"
        f"💰 Amount: *SGD {amount:,.2f}*\n\n"
        "👤 *Who paid?*",
        reply_markup=_payer_keyboard(),
        parse_mode="Markdown",
    )
    return SELECTING_PAYER


# ── Step 4 — Payer selected → ask for date ────────────────────────────────────

async def step_payer(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()

    payer = query.data.split("payer:", 1)[1]
    context.user_data["payer"] = payer

    entity = context.user_data["entity"]
    amount = context.user_data["amount"]
    today  = datetime.now()

    await query.edit_message_text(
        f"📁 Entity: *{entity}*\n"
        f"💰 Amount: *SGD {amount:,.2f}*\n"
        f"👤 Payer: *{payer}*\n\n"
        "📅 *Date of expense?*\n"
        "_Tap the button to use today, or type a date (e.g. `25 Jun` or `25/06/2026`)_",
        reply_markup=_date_keyboard(today),
        parse_mode="Markdown",
    )
    return ENTERING_DATE


# ── Step 5a — Date button (today) ────────────────────────────────────────────

async def step_date_today(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()

    today = datetime.now()
    context.user_data["expense_date"] = today

    entity = context.user_data["entity"]
    amount = context.user_data["amount"]
    payer  = context.user_data["payer"]

    await query.edit_message_text(
        f"📁 Entity: *{entity}*\n"
        f"💰 Amount: *SGD {amount:,.2f}*\n"
        f"👤 Payer: *{payer}*\n"
        f"📅 Date: *{today.strftime('%-d %b %Y')}*\n\n"
        "📝 *What is this expense for? (purpose / description)*",
        parse_mode="Markdown",
    )
    return ENTERING_PURPOSE


# ── Step 5b — Date typed by user ─────────────────────────────────────────────

async def step_date_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    dt = _parse_date(update.message.text)
    if dt is None:
        await update.message.reply_text(
            "❌ Couldn't read that date. Try `27 Jun`, `27/06/2026`, or tap the button.",
            reply_markup=_date_keyboard(datetime.now()),
            parse_mode="Markdown",
        )
        return ENTERING_DATE

    context.user_data["expense_date"] = dt

    entity = context.user_data["entity"]
    amount = context.user_data["amount"]
    payer  = context.user_data["payer"]

    await update.message.reply_text(
        f"📁 Entity: *{entity}*\n"
        f"💰 Amount: *SGD {amount:,.2f}*\n"
        f"👤 Payer: *{payer}*\n"
        f"📅 Date: *{dt.strftime('%-d %b %Y')}*\n\n"
        "📝 *What is this expense for? (purpose / description)*",
        parse_mode="Markdown",
    )
    return ENTERING_PURPOSE


# ── Step 6 — Purpose entered → save everything ───────────────────────────────

async def step_purpose(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    purpose = update.message.text.strip()
    if not purpose:
        await update.message.reply_text("❌ Please enter a description.")
        return ENTERING_PURPOSE

    context.user_data["purpose"] = purpose

    entity       = context.user_data["entity"]
    amount       = context.user_data["amount"]
    payer        = context.user_data["payer"]
    expense_date = context.user_data["expense_date"]
    image_bytes  = context.user_data["image_bytes"]

    status_msg = await update.message.reply_text("⏳ Saving expense…")

    try:
        # ── Build receipt filename ────────────────────────────────────────────
        filename  = _receipt_filename(expense_date, purpose)
        folder    = ENTITY_FOLDER_NAMES.get(entity, entity)
        receipt_folder = f"{RECEIPTS_BASE_PATH}/{folder}"
        receipt_path   = f"{receipt_folder}/{filename}"

        # ── Upload receipt image ──────────────────────────────────────────────
        await status_msg.edit_text("⏳ Uploading receipt to OneDrive…")
        await asyncio.get_event_loop().run_in_executor(
            None, graph.ensure_folder, receipt_folder
        )
        upload_result = await asyncio.get_event_loop().run_in_executor(
            None,
            lambda: graph.upload_file(receipt_path, image_bytes, "image/jpeg"),
        )
        receipt_url = upload_result.get("webUrl", "")

        # ── Download, update, re-upload Excel ────────────────────────────────
        await status_msg.edit_text("⏳ Updating expense tracker…")
        excel_bytes = await asyncio.get_event_loop().run_in_executor(
            None, lambda: graph.download_file(EXCEL_FILE_PATH)
        )
        if excel_bytes is None:
            await status_msg.edit_text(
                f"❌ Excel file not found at `{EXCEL_FILE_PATH}`.",
                parse_mode="Markdown",
            )
            return ConversationHandler.END

        updated = await asyncio.get_event_loop().run_in_executor(
            None,
            lambda: add_expense(
                excel_bytes,
                sheet_name=entity,
                purpose=purpose,
                amount=amount,
                expense_date=expense_date,
                payer=payer,
                receipt_filename=filename,
            ),
        )
        await asyncio.get_event_loop().run_in_executor(
            None,
            lambda: graph.upload_file(
                EXCEL_FILE_PATH,
                updated,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ),
        )

        # ── Confirmation message ──────────────────────────────────────────────
        full_name  = ENTITY_FULL_NAMES.get(entity, entity)
        link_line  = f"\n🖼 [View receipt]({receipt_url})" if receipt_url else ""
        await status_msg.edit_text(
            f"✅ *Expense saved!*\n\n"
            f"🏢 {full_name}\n"
            f"📝 {purpose}\n"
            f"💰 SGD {amount:,.2f}\n"
            f"👤 {payer}\n"
            f"📅 {expense_date.strftime('%-d %b %Y')}"
            + link_line,
            parse_mode="Ma