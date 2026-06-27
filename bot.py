import asyncio, logging, re, unicodedata, io
from datetime import datetime
import msal, requests, openpyxl
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CallbackQueryHandler, CommandHandler, ConversationHandler, MessageHandler, filters, ContextTypes
import os
from dotenv import load_dotenv

load_dotenv()
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Config
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

# Graph client
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
_msal_app = msal.ConfidentialClientApplication(
    client_id=AZURE_CLIENT_ID,
    client_credential=AZURE_CLIENT_SECRET,
    authority=f"https://login.microsoftonline.com/{AZURE_TENANT_ID}",
)
_site_id_cache = None
_drive_id_cache = None

def _token():
    r = _msal_app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in r:
        raise RuntimeError(r.get("error_description"))
    return r["access_token"]

def _h():
    return {"Authorization": f"Bearer {_token()}"}

def _site_id():
    global _site_id_cache
    if not _site_id_cache:
        r = requests.get(f"{GRAPH_BASE}/sites/{SHAREPOINT_HOSTNAME}:{SHAREPOINT_SITE_PATH}", headers=_h())
        r.raise_for_status()
        _site_id_cache = r.json()["id"]
    return _site_id_cache

def _drive_id():
    global _drive_id_cache
    if not _drive_id_cache:
        r = requests.get(f"{GRAPH_BASE}/sites/{_site_id()}/drives", headers=_h())
        r.raise_for_status()
        drives = r.json().get("value", [])
        for d in drives:
            if d.get("name","").lower() == SHAREPOINT_DRIVE_NAME.lower():
