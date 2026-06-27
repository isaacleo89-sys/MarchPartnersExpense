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

# Conversation states
(
    SELECTING_ENTITY,
    ENTERING_AMOUNT,
    SELECTING_PAYER,
    ENTERING_DATE,
    ENTERING_PURPOSE,
) = range(5)

graph = GraphClient()


def _is_authorised(user_id: int) -> bool:
    return not AUTHORIZED_USER_IDS or user_id in AUTHORIZED_USER_IDS


def _entity_keyboard() -> InlineKeyboardMarkup:
    sheets = list(dict.fromkeys(ENTITY_MAP.values()))
    buttons = [InlineKeyboardButton(name, callback_data=f"entity:{name}") for name in sheets]
    rows = [buttons[i:i + 2] for i in range(0, len(buttons), 2)]
    return InlineKeyboardMarkup(rows)


def _payer_keyboard() -> InlineKeyboardMarkup:
    buttons = [InlineKeyboardButton(p, callback_data=f"payer:{p}") for p in PAYERS]
    return InlineKeyboardMarkup([buttons])


def _date_keyboard(today: datetime) -> InlineKeyboardMarkup:
    label = "Today - " + today.strftime("%d %b %Y")
    return InlineKeyboardMarkup([[InlineKeyboardButton(label, callback_data="date:TODAY")]])


def _receipt_filename(expense_date: datetime, purpose: str) -> str:
    date_str = expense_date.strftime("%Y%b%d")
    safe = unicodedata.normalize("NFKD", purpose).encode("ascii", "ignore").decode()
    safe = re.sub(r"[^\w\s-]", "", safe).strip()
    safe = re.sub(r"\s+", "_", safe)[:60]
    return f"{date_str}_{safe}.jpg"


_DATE_FORMATS = [
    "%d %b %Y", "%d %B %Y",
    "%d %b", "%d %B",
    "%d/%m/%Y", "%d-%m-%Y",
    "%Y-%m-%d",
]

def _parse_date(text: str) -> datetime | None:
    text = text.strip()
    for fmt in _DATE_FORMATS:
        try:
            dt = datetime.strptime(text, fmt)
            if dt.year == 1900:
                dt = dt.replace(year=datetime.now().year)
            return dt
        except ValueError:
            continue
    return None


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not _is_authorised(update.effective_user.id):
        return
    await update.message.reply_text(
        "March Partners Expense Bot\n\n"
        "Send me a photo of your receipt and I'll guide you through the rest.\n\n"
        "/cancel - cancel the current entry"
    )


async def cmd_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    await update.message.reply_text("Entry cancelled.")
    return ConversationHandler.END


async def step_photo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not _is_authorised(update.effective_user.id):
        return ConversationHandler.END
    photo = update.message.photo[-1]
    photo_file = await photo.get_file()
    image_bytes = bytes(await photo_file.download_as_bytearray())
    context.user_data.clear()
    context.user_data["image_bytes"] = image_bytes
    await update.message.reply_text(
        "Which entity should this be charged to?",
        reply_markup=_entity_keyboard(),
    )
    return SELECTING_ENTITY


async def step_entity(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()
    sheet_name = query.data.split("entity:", 1)[1]
    context.user_data["entity"] = sheet_name
    await query.edit_message_text(
        f"Entity: {sheet_name}\n\nEnter the amount (SGD):"
    )
    return ENTERING_AMOUNT


async def step_amount(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    text = update.message.text.strip()
    cleaned = re.sub(r"[^\d.]", "", text)
    try:
        amount = float(cleaned)
        if amount <= 0:
            raise ValueError
    except ValueError:
        await update.message.reply_text("Please enter a valid amount, e.g. 123.45")
        return ENTERING_AMOUNT
    context.user_data["amount"] = amount
    entity = context.user_data["entity"]
    await update.message.reply_text(
        f"Entity: {entity}\nAmount: SGD {amount:,.2f}\n\nWho paid?",
        reply_markup=_payer_keyboard(),
    )
    return SELECTING_PAYER


async def step_payer(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()
    payer = query.data.split("payer:", 1)[1]
    context.user_data["payer"] = payer
    entity = context.user_data["entity"]
    amount = context.user_data["amount"]
    today = datetime.now()
    await query.edit_message_text(
        f"Entity: {entity}\nAmount: SGD {amount:,.2f}\nPayer: {payer}\n\n"
        "Date of expense? Tap the button for today, or type a date (e.g. 25 Jun or 25/06/2026)",
        reply_markup=_date_keyboard(today),
    )
    return ENTERING_DATE


async def step_date_today(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()
    today = datetime.now()
    context.user_data["expense_date"] = today
    entity = context.user_data["entity"]
    amount = context.user_data["amount"]
    payer = context.user_data["payer"]
    await query.edit_message_text(
        f"Entity: {entity}\nAmount: SGD {amount:,.2f}\nPayer: {payer}\n"
        f"Date: {today.strftime('%d %b %Y')}\n\nWhat is this expense for?"
    )
    return ENTERING_PURPOSE


async def step_date_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    dt = _parse_date(update.message.text)
    if dt is None:
        await update.message.reply_text(
            "Could not read that date. Try: 27 Jun, 27/06/2026, or tap the button.",
            reply_markup=_date_keyboard(datetime.now()),
        )
        return ENTERING_DATE
    context.user_data["expense_date"] = dt
    entity = context.user_data["entity"]
    amount = context.user_data["amount"]
    payer = context.user_data["payer"]
    await update.message.reply_text(
        f"Entity: {entity}\nAmount: SGD {amount:,.2f}\nPayer: {payer}\n"
        f"Date: {dt.strftime('%d %b %Y')}\n\nWhat is this expense for?"
    )
    return ENTERING_PURPOSE


async def step_purpose(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    purpose = update.message.text.strip()
    if not purpose:
        await update.message.reply_text("Please enter a description.")
        return ENTERING_PURPOSE
    context.user_data["purpose"] = purpose
    entity = context.user_data["entity"]
    amount = context.user_data["amount"]
    payer = context.user_data["payer"]
    expense_date = context.user_data["expense_date"]
    image_bytes = context.user_data["image_bytes"]
    status_msg = await update.message.reply_text("Saving expense...")
    try:
        filename = _receipt_filename(expense_date, purpose)
        folder = ENTITY_FOLDER_NAMES.get(entity, entity)
        receipt_folder = f"{RECEIPTS_BASE_PATH}/{folder}"
        receipt_path = f"{receipt_folder}/{filename}"
        await status_msg.edit_text("Uploading receipt to OneDrive...")
        await asyncio.get_event_loop().run_in_executor(None, graph.ensure_folder, receipt_folder)
        upload_result = await asyncio.get_event_loop().run_in_executor(
            None, lambda: graph.upload_file(receipt_path, image_bytes, "image/jpeg")
        )
        receipt_url = upload_result.get("webUrl", "")
        await status_msg.edit_text("Updating expense tracker...")
        excel_bytes = await asyncio.get_event_loop().run_in_executor(
            None, lambda: graph.download_file(EXCEL_FILE_PATH)
        )
        if excel_bytes is None:
            await status_msg.edit_text(f"Excel file not found at {EXCEL_FILE_PATH}")
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
        full_name = ENTITY_FULL_NAMES.get(entity, entity)
        confirm = (
            f"Expense saved!\n\n"
            f"{full_name}\n"
            f"{purpose}\n"
            f"SGD {amount:,.2f}\n"
            f"{payer}\n"
            f"{expense_date.strftime('%d %b %Y')}"
        )
        if receipt_url:
            confirm += f"\nReceipt: {receipt_url}"
        await status_msg.edit_text(confirm)
    except Exception as e:
        logger.exception("Error saving expense")
        await status_msg.edit_text(f"Error: {e}")
    finally:
        context.user_data.clear()
    return ConversationHandler.END


def main() -> None:
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    conv = ConversationHandler(
        entry_points=[MessageHandler(filters.PHOTO, step_photo)],
        states={
            SELECTING_ENTITY: [CallbackQueryHandler(step_entity, pattern=r"^entity:")],
            ENTERING_AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, step_amount)],
            SELECTING_PAYER: [CallbackQueryHandler(step_payer, pattern=r"^payer:")],
            ENTERING_DATE: [
                CallbackQueryHandler(step_date_today, pattern=r"^date:TODAY$"),
                MessageHandler(filters.TEXT & ~filters.COMMAND, step_date_text),
            ],
            ENTERING_PURPOSE: [MessageHandler(filters.TEXT & ~filters.COMMAND, step_purpose)],
        },
        fallbacks=[CommandHandler("cancel", cmd_cancel)],
        per_user=True,
        per_chat=True,
    )
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("help", cmd_start))
    app.add_handler(conv)
    logger.info("Bot is running...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
