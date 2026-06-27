# Expense Bot — Setup Guide

## What this bot does

Send a receipt photo to Telegram → bot extracts details → receipt saved to OneDrive → row appended to `March Partners Entity Expenses.xlsx`.

---

## Step 1 — Create the Telegram bot

1. Open Telegram and message **@BotFather**
2. Send `/newbot`, give it a name and username (e.g. `MarchExpensesBot`)
3. Copy the **token** (looks like `123456:ABC-DEF...`)

To find your own Telegram user ID, message **@userinfobot**.

---

## Step 2 — Create an Azure App Registration

This gives the bot permission to read and write files in your SharePoint.

1. Go to [portal.azure.com](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**
2. Name it something like `ExpenseBot`, leave everything else as default, click **Register**
3. Note the **Application (client) ID** and **Directory (tenant) ID** from the overview page
4. Go to **Certificates & secrets** → **New client secret** → copy the **Value** immediately
5. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions**
6. Search for and add: **`Sites.ReadWrite.All`**
7. Click **Grant admin consent** (you need to be a Global Admin or have an admin do this)

---

## Step 3 — Find your SharePoint hostname and site path

Your OneDrive for Business syncs from SharePoint. To find the details:

1. Open any file in the `Expenses` folder via the OneDrive web interface
2. Look at the URL — it will look like:
   `https://marchpartners.sharepoint.com/sites/MarchPartners/Shared Documents/...`
3. From that URL:
   - `SHAREPOINT_HOSTNAME` = `marchpartners.sharepoint.com`
   - `SHAREPOINT_SITE_PATH` = `/sites/MarchPartners` (or `/` if it's the root)
   - `SHAREPOINT_DRIVE_NAME` = the library name, usually `Documents`

---

## Step 4 — Configure the bot

1. Copy `.env.example` to `.env`
2. Fill in all values

```
TELEGRAM_TOKEN=...
AUTHORIZED_USER_IDS=your-telegram-user-id
AZURE_TENANT_ID=...
AZURE_CLIENT_ID=...
AZURE_CLIENT_SECRET=...
SHAREPOINT_HOSTNAME=marchpartners.sharepoint.com
SHAREPOINT_SITE_PATH=/sites/MarchPartners
SHAREPOINT_DRIVE_NAME=Documents
EXCEL_FILE_PATH=Expenses/March Partners Entity Expenses.xlsx
RECEIPTS_BASE_PATH=Expenses/Receipts
```

---

## Step 5 — Deploy to a cloud server

### Option A — Railway (easiest, free tier available)

1. Push the `expense-bot` folder to a GitHub repo
2. Go to [railway.app](https://railway.app) → New Project → Deploy from GitHub
3. Add all your `.env` variables under **Variables**
4. Railway will build the Docker image and run the bot automatically

### Option B — Any VPS (DigitalOcean, Vultr, Linode, etc.)

```bash
git clone https://github.com/you/expense-bot.git
cd expense-bot
cp .env.example .env
# edit .env with your values
docker build -t expense-bot .
docker run -d --env-file .env --restart unless-stopped expense-bot
```

---

## Using the bot

Send a receipt **photo** with this caption:

```
entity | item | category | amount | payer [| offset [| comment]]
```

**Entity shortcuts:**

| Shortcut | Entity |
|----------|--------|
| `march` or `gp` | March Real Estate Partners Pte Ltd (GP) |
| `heritage` or `fund` | Heritage Strata Commercial I Pte Ltd (Fund Co) |
| `falcon` or `61ecr` | Falcon Prosperity Pte Ltd (61 ECR) |
| `effraie` or `bs` | Effraie Prosperity Pte Ltd (Boon Sing) |
| `crystal` or `geylang` | Crystal Pinnacle Pte Ltd (253A Geylang) |
| `crimson` or `binjai` | Crimson Phoenix Pte Ltd (25 Binjai Park) |

**Examples:**

```
falcon | SPA legal fees | Legal | 1500 | IL | No
march | Coffee with Gabriel | Entertainment | 33.58 | JK, IL | N/A
gp | Lunch with Savills | Entertainment | | JK, IL | N/A | Sophia meeting
```

- **Amount** can be left blank — the bot will try OCR, then ask you
- **Offset** options: `Yes` / `No` / `N/A` (defaults to `No`)
- The receipt image is saved to `Expenses/Receipts/{Entity}/` in SharePoint
- The filename is added to the **Comment** column so you can trace it

---

## Commands

| Command | Description |
|---------|-------------|
| `/start` or `/help` | Show usage instructions |
| `/entities` | List all entity shortcuts |
