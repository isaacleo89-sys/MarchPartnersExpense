import asyncio, logging, re, unicodedata, io, os
from datetime import datetime
from dotenv import load_dotenv
import msal, requests, openpyxl
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (Application, CallbackQueryHandler, CommandHandler,
    ConversationHandler, MessageHandler, filters, ContextTypes)

load_dotenv()
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

TELEGRAM_TOKEN = os.environ["TELEGRAM_TOKEN"]
AUTHORIZED_USER_IDS = [int(x) for x in os.getenv("AUTHORIZED_USER_IDS","").split(",") if x.strip()]
AZURE_TENANT_ID = os.environ["AZURE_TENANT_ID"]
AZURE_CLIENT_ID = os.environ["AZURE_CLIENT_ID"]
AZURE_CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
SHAREPOINT_HOSTNAME = os.environ["SHAREPOINT_HOSTNAME"]
SHAREPOINT_SITE_PATH = os.getenv("SHAREPOINT_SITE_PATH", "/")
SHAREPOINT_DRIVE_NAME = os.getenv("SHAREPOINT_DRIVE_NAME", "Documents")
EXCEL_FILE_PATH = os.getenv("EXCEL_FILE_PATH", "Expenses/March Partners Entity Expenses.xlsx")
RECEIPTS_BASE_PATH = os.getenv("RECEIPTS_BASE_PATH", "Expenses")

PAYERS = ["IL", "JK", "Fund"]
ENTITY_MAP = {
    "march": "March (GP)", "gp": "March (GP)",
    "heritage": "Heritage (Fund)", "fund": "Heritage (Fund)",
    "falcon": "Falcon (61ECR)", "61ecr": "Falcon (61ECR)",
    "effraie": "Effraie (BS)", "bs": "Effraie (BS)",
    "crystal": "Crystal (Geylang)", "geylang": "Crystal (Geylang)",
    "crimson": "Crimson (Binjai)", "binjai": "Crimson (Binjai)",
}
ENTITY_FULL_NAMES = {
    "March (GP)": "March Real Estate Partners Pte Ltd (GP)",
    "Heritage (Fund)": "Heritage Strata Commercial I Pte Ltd (Fund Co)",
    "Falcon (61ECR)": "Falcon Prosperity Pte Ltd (61 ECR)",
    "Effraie (BS)": "Effraie Prosperity Pte Ltd (Boon Sing)",
    "Crystal (Geylang)": "Crystal Pinnacle Pte Ltd (253A Geylang)",
    "Crimson (Binjai)": "Crimson Phoenix Pte Ltd (25 Binjai Park)",
}
ENTITY_FOLDER_NAMES = {
    "March (GP)": "March", "Heritage (Fund)": "Heritage",
    "Falcon (61ECR)": "Falcon", "Effraie (BS)": "Effraie",
    "Crystal (Geylang)": "Crystal", "Crimson (Binjai)": "Crimson",
}

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
_msal_app = msal.ConfidentialClientApplication(
    client_id=AZURE_CLIENT_ID,
    client_credential=AZURE_CLIENT_SECRET,
    authority="https://login.microsoftonline.com/" + AZURE_TENANT_ID,
)
_site_id_cache = None
_drive_id_cache = None


def _token():
    r = _msal_app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in r:
        raise RuntimeError(r.get("error_description"))
    return r["access_token"]


def _h():
    return {"Authorization": "Bearer " + _token()}


def _site_id():
    global _site_id_cache
    if not _site_id_cache:
        r = requests.get(GRAPH_BASE + "/sites/" + SHAREPOINT_HOSTNAME + ":" + SHAREPOINT_SITE_PATH, headers=_h())
        r.raise_for_status()
        _site_id_cache = r.json()["id"]
    return _site_id_cache


def _drive_id():
    global _drive_id_cache
    if not _drive_id_cache:
        r = requests.get(GRAPH_BASE + "/sites/" + _site_id() + "/drives", headers=_h())
        r.raise_for_status()
        drives = r.json().get("value", [])
        for d in drives:
            if d.get("name", "").lower() == SHAREPOINT_DRIVE_NAME.lower():
                _drive_id_cache = d["id"]
                break
        if not _drive_id_cache and drives:
            _drive_id_cache = drives[0]["id"]
    return _drive_id_cache


def _item_url(path):
    return GRAPH_BASE + "/drives/" + _drive_id() + "/root:/" + path.strip("/")


def download_file(path):
    r = requests.get(_item_url(path) + ":/content", headers=_h())
    if r.status_code == 404:
        return None
    r.raise_for_status()
    return r.content


def upload_file(path, content, ct="application/octet-stream"):
    r = requests.put(_item_url(path) + ":/content", headers=dict(list(_h().items()) + [("Content-Type", ct)]), data=content)
    r.raise_for_status()
    return r.json()


def ensure_folder(path):
    parts = path.strip("/").split("/")
    for depth in range(1, len(parts) + 1):
        p = parts[:depth]
        parent = "/".join(p[:-1])
        name = p[-1]
        if parent:
            url = GRAPH_BASE + "/drives/" + _drive_id() + "/root:/" + parent + ":/children"
        else:
            url = GRAPH_BASE + "/drives/" + _drive_id() + "/root/children"
        try:
            r = requests.post(url, headers=dict(list(_h().items()) + [("Content-Type", "application/json")]),
                json={"name": name, "folder": {}, "@microsoft.graph.conflictBehavior": "replace"})
            r.raise_for_status()
        except requests.HTTPError as e:
            if e.response is not None and e.response.status_code in (409, 403):
                pass
            else:
                raise


def add_expense(wb_bytes, sheet_name, purpose, amount, expense_date, payer, receipt_filename):
    wb = openpyxl.load_workbook(io.BytesIO(wb_bytes))
    if sheet_name not in wb.sheetnames:
        raise ValueError("Sheet not found: " + sheet_name)
    ws = wb[sheet_name]
    last = 4
    for row in ws.iter_rows(min_row=5):
        if any(c.value is not None for c in row):
            last = row[0].row
    new = last + 1
    ws.cell(new, 2).value = 1 if last == 4 else ("=B" + str(last) + "+1")
    ws.cell(new, 3).value = purpose
    ws.cell(new, 5).value = amount
    ws.cell(new, 5).number_format = "#,##0.00"
    ws.cell(new, 6).value = expense_date
    ws.cell(new, 6).number_format = "D MMM YYYY"
    ws.cell(new, 7).value = payer
    ws.cell(new, 8).value = "N/A"
    ws.cell(new, 9).value = receipt_filename
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


SELECTING_ENTITY, ENTERING_AMOUNT, SELECTING_PAYER, ENTERING_DATE, ENTERING_PURPOSE = range(5)


def _auth(uid):
    return not AUTHORIZED_USER_IDS or uid in AUTHORIZED_USER_IDS


def _entity_kb():
    sheets = list(dict.fromkeys(ENTITY_MAP.values()))
    btns = [InlineKeyboardButton(s, callback_data="entity:" + s) for s in sheets]
    return InlineKeyboardMarkup([btns[i:i+2] for i in range(0, len(btns), 2)])


def _payer_kb():
    return InlineKeyboardMarkup([[InlineKeyboardButton(p, callback_data="payer:" + p) for p in PAYERS]])


def _date_kb(today):
    return InlineKeyboardMarkup([[InlineKeyboardButton("Today - " + today.strftime("%d %b %Y"), callback_data="date:TODAY")]])


def _filename(dt, purpose):
    date_str = dt.strftime("%Y%b%d")
    safe = unicodedata.normalize("NFKD", purpose).encode("ascii", "ignore").decode()
    safe = re.sub(r"[^\w\s-]", "", safe).strip()
    safe = re.sub(r"\s+", "_", safe)[:60]
    return date_str + "_" + safe + ".jpg"


_DATE_FMTS = ["%d %b %Y", "%d %B %Y", "%d %b", "%d %B", "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"]


def _parse_date(text):
    for fmt in _DATE_FMTS:
        try:
            dt = datetime.strptime(text.strip(), fmt)
            return dt.replace(year=datetime.now().year) if dt.year == 1900 else dt
        except ValueError:
            pass
    return None


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not _auth(update.effective_user.id):
        return
    await update.message.reply_text("March Partners Expense Bot\n\nSend me a receipt photo to get started.\n\n/cancel - cancel current entry")


async def cmd_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("Entry cancelled.")
    return ConversationHandler.END


async def step_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not _auth(update.effective_user.id):
        return ConversationHandler.END
    pf = await update.message.photo[-1].get_file()
    context.user_data.clear()
    context.user_data["image_bytes"] = bytes(await pf.download_as_bytearray())
    await update.message.reply_text("Which entity?", reply_markup=_entity_kb())
    return SELECTING_ENTITY


async def step_entity(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    sheet = q.data.split("entity:", 1)[1]
    context.user_data["entity"] = sheet
    await q.edit_message_text("Entity: " + sheet + "\n\nEnter the amount (SGD):")
    return ENTERING_AMOUNT


async def step_amount(update: Update, context: ContextTypes.DEFAULT_TYPE):
    cleaned = re.sub(r"[^\d.]", "", update.message.text.strip())
    try:
        amount = float(cleaned)
        if amount <= 0:
            raise ValueError
    except (ValueError, TypeError):
        await update.message.reply_text("Please enter a valid amount e.g. 123.45")
        return ENTERING_AMOUNT
    context.user_data["amount"] = amount
    entity = context.user_data["entity"]
    await update.message.reply_text(
        "Entity: " + entity + "\nAmount: SGD " + f"{amount:,.2f}" + "\n\nWho paid?",
        reply_markup=_payer_kb())
    return SELECTING_PAYER


async def step_payer(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    payer = q.data.split("payer:", 1)[1]
    context.user_data["payer"] = payer
    entity = context.user_data["entity"]
    amount = context.user_data["amount"]
    today = datetime.now()
    await q.edit_message_text(
        "Entity: " + entity + "\nAmount: SGD " + f"{amount:,.2f}" + "\nPayer: " + payer + "\n\nDate? (tap for today or type e.g. 25 Jun)",
        reply_markup=_date_kb(today))
    return ENTERING_DATE


async def step_date_today(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    today = datetime.now()
    context.user_data["expense_date"] = today
    entity = context.user_data["entity"]
    amount = context.user_data["amount"]
    payer = context.user_data["payer"]
    await q.edit_message_text(
        "Entity: " + entity + "\nAmount: SGD " + f"{amount:,.2f}" + "\nPayer: " + payer +
        "\nDate: " + today.strftime("%d %b %Y") + "\n\nWhat is this expense for?")
    return ENTERING_PURPOSE


async def step_date_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    dt = _parse_date(update.message.text)
    if not dt:
        await update.message.reply_text("Could not read date. Try: 27 Jun or 27/06/2026", reply_markup=_date_kb(datetime.now()))
        return ENTERING_DATE
    context.user_data["expense_date"] = dt
    entity = context.user_data["entity"]
    amount = context.user_data["amount"]
    payer = context.user_data["payer"]
    await update.message.reply_text(
        "Entity: " + entity + "\nAmount: SGD " + f"{amount:,.2f}" + "\nPayer: " + payer +
        "\nDate: " + dt.strftime("%d %b %Y") + "\n\nWhat is this expense for?")
    return ENTERING_PURPOSE


async def step_purpose(update: Update, context: ContextTypes.DEFAULT_TYPE):
    purpose = update.message.text.strip()
    if not purpose:
        await update.message.reply_text("Please enter a description.")
        return ENTERING_PURPOSE
    entity = context.user_data["entity"]
    amount = context.user_data["amount"]
    payer = context.user_data["payer"]
    expense_date = context.user_data["expense_date"]
    image_bytes = context.user_data["image_bytes"]
    status = await update.message.reply_text("Saving...")
    try:
        filename = _filename(expense_date, purpose)
        folder = ENTITY_FOLDER_NAMES.get(entity, entity)
        receipt_folder = RECEIPTS_BASE_PATH + "/" + folder
        receipt_path = receipt_folder + "/" + filename
        await status.edit_text("Uploading receipt...")
        await asyncio.get_event_loop().run_in_executor(None, ensure_folder, receipt_folder)
        result = await asyncio.get_event_loop().run_in_executor(None, lambda: upload_file(receipt_path, image_bytes, "image/jpeg"))
        receipt_url = result.get("webUrl", "")
        await status.edit_text("Updating spreadsheet...")
        wb_bytes = await asyncio.get_event_loop().run_in_executor(None, lambda: download_file(EXCEL_FILE_PATH))
        if not wb_bytes:
            await status.edit_text("Excel not found at " + EXCEL_FILE_PATH)
            return ConversationHandler.END
        updated = await asyncio.get_event_loop().run_in_executor(None, lambda: add_expense(wb_bytes, entity, purpose, amount, expense_date, payer, filename))
        await asyncio.get_event_loop().run_in_executor(None, lambda: upload_file(EXCEL_FILE_PATH, updated, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
        full_name = ENTITY_FULL_NAMES.get(entity, entity)
        msg = "Saved!\n\n" + full_name + "\n" + purpose + "\nSGD " + f"{amount:,.2f}" + "\n" + payer + "\n" + expense_date.strftime("%d %b %Y")
        if receipt_url:
            msg += "\n" + receipt_url
        await status.edit_text(msg)
    except Exception as e:
        logger.exception("Error")
        await status.edit_text("Error: " + str(e))
    finally:
        context.user_data.clear()
    return ConversationHandler.END


def main():
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
    logger.info("Running...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
