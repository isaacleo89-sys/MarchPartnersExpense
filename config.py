import os
from dotenv import load_dotenv

load_dotenv()

# ── Telegram ──────────────────────────────────────────────────────────────────
TELEGRAM_TOKEN = os.environ["TELEGRAM_TOKEN"]

# Comma-separated Telegram user IDs allowed to use the bot (leave empty to allow all)
AUTHORIZED_USER_IDS: list[int] = [
    int(x) for x in os.getenv("AUTHORIZED_USER_IDS", "").split(",") if x.strip()
]

# ── Microsoft Azure / Graph API ───────────────────────────────────────────────
AZURE_TENANT_ID = os.environ["AZURE_TENANT_ID"]
AZURE_CLIENT_ID = os.environ["AZURE_CLIENT_ID"]
AZURE_CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]

# SharePoint site details
# SHAREPOINT_HOSTNAME: e.g. "marchpartners.sharepoint.com"
# SHAREPOINT_SITE_PATH: e.g. "/" for root or "/sites/MarchPartners"
SHAREPOINT_HOSTNAME = os.environ["SHAREPOINT_HOSTNAME"]
SHAREPOINT_SITE_PATH = os.getenv("SHAREPOINT_SITE_PATH", "/")
SHAREPOINT_DRIVE_NAME = os.getenv("SHAREPOINT_DRIVE_NAME", "Documents")

# ── File paths (relative to SharePoint drive root) ───────────────────────────
EXCEL_FILE_PATH = os.getenv(
    "EXCEL_FILE_PATH",
    "Expenses/March Partners Entity Expenses.xlsx",
)
RECEIPTS_BASE_PATH = os.getenv("RECEIPTS_BASE_PATH", "Expenses")

# ── Payer options ─────────────────────────────────────────────────────────────
PAYERS: list[str] = ["IL", "JK", "Fund"]

# ── Entity mapping ────────────────────────────────────────────────────────────
# Keys are what the user types (case-insensitive); values are exact sheet tab names.
ENTITY_MAP: dict[str, str] = {
    "march": "March (GP)",
    "gp": "March (GP)",
    "heritage": "Heritage (Fund)",
    "fund": "Heritage (Fund)",
    "falcon": "Falcon (61ECR)",
    "61ecr": "Falcon (61ECR)",
    "effraie": "Effraie (BS)",
    "bs": "Effraie (BS)",
    "crystal": "Crystal (Geylang)",
    "geylang": "Crystal (Geylang)",
    "crimson": "Crimson (Binjai)",
    "binjai": "Crimson (Binjai)",
}

# Full legal names (for confirmation messages)
ENTITY_FULL_NAMES: dict[str, str] = {
    "March (GP)": "March Real Estate Partners Pte Ltd (GP)",
    "Heritage (Fund)": "Heritage Strata Commercial I Pte Ltd (Fund Co)",
    "Falcon (61ECR)": "Falcon Prosperity Pte Ltd (61 ECR)",
    "Effraie (BS)": "Effraie Prosperity Pte Ltd (Boon Sing)",
    "Crystal (Geylang)": "Crystal Pinnacle Pte Ltd (253A Geylang)",
    "Crimson (Binjai)": "Crimson Phoenix Pte Ltd (25 Binjai Park)",
}

# Short folder names used when saving receipt image