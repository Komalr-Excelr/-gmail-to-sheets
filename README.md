# Gmail to Google Sheets (Python)

Reads unread Gmail messages and appends Sender, Subject, Date, and Body as new rows in a Google Sheet. After logging, emails are marked as read. State is stored locally to avoid duplicates across runs.

## Architecture

┌──────────────────────────────────────────────┐
│        GMAIL → GOOGLE SHEETS SYSTEM          │
└──────────────────────────────────────────────┘

               ┌──────────────────────────────┐
               │  USER PC / LAPTOP            │
               │ (Runs Python Script)         │
               └───────────────┬──────────────┘
                               │
                               ▼

┌──────────────────────────────────────────────┐
│      LOCAL PROJECT SETUP (gmail-to-sheets)   │
│                                              │
│ main.py     → Script logic                   │
│ token.json  → Saved OAuth login              │
│ credentials/credentials.json  → OAuth ID     │
└──────────────────────────────────────────────┘
                               │
                               ▼

┌──────────────────────────────────────────────┐
│     GOOGLE AUTHENTICATION (OAuth Flow)       │
│                                              │
│ credentials.json → Opens browser login       │
│ User allows Gmail + Sheets access            │
│ token.json created for future runs           │
└──────────────────────────────────────────────┘
                               │
                               ▼

┌──────────────────────────────────────────────┐
│           GMAIL MESSAGE PROCESSING           │
│                                              │
│ Fetch unread messages (Gmail API)            │
│ Extract: date, from, subject, body           │
│ Mark email as READ                           │
└──────────────────────────────────────────────┘
                               │
                               ▼

┌──────────────────────────────────────────────┐
│             GOOGLE SHEETS UPDATE             │
│                                              │
│ Sheets API appends rows                      │
│ Spreadsheet ID taken from URL                │
│ Stored columns: Date | From | Subject | Body │
└──────────────────────────────────────────────┘
                               │
                               ▼

┌──────────────────────────────────────────────┐
│                  OUTPUT                       │
│                                              │
│ ✔ Emails saved to Sheet                      │
│ ✔ No duplicates (already marked read)        │
│ ✔ Run again anytime                          │
└──────────────────────────────────────────────┘


## What it does
- Authenticates via OAuth 2.0 (no service accounts).
- Fetches unread emails from your Gmail inbox.
- Extracts Sender, Subject, Date, Body.
- Appends new rows to a Google Sheet.
- Marks emails as read after successful logging.
- Prevents duplicates by tracking processed message IDs across runs.

## Project Structure
```
gmail-to-sheets/
├── src/
│   ├── gmail_service.py   # Gmail API interactions
│   ├── sheets_service.py  # Sheets API interactions
│   ├── email_parser.py    # Parse email content
│   └── main.py            # Orchestrator
├── credentials/
│   └── credentials.json   # OAuth credentials (DO NOT COMMIT)
├── .gitignore
├── requirements.txt
├── README.md
└── config.py
```

## Prerequisites
- Python 3.8+
- A Google account with Gmail and Sheets access
- A Google Cloud project with Gmail API and Google Sheets API enabled

## Setup (Windows)
1. Enable APIs:
   - Open https://console.cloud.google.com/apis/dashboard
   - Create/select a project.
   - Enable: Gmail API and Google Sheets API.

2. Create OAuth credentials (Installed App):
   - In Google Cloud Console, go to "APIs & Services" → "Credentials" → "Create Credentials" → "OAuth client ID".
   - Application type: "Desktop app".
   - Download the JSON and save it as `credentials/credentials.json` in this project.

3. Create a Google Sheet:
   - Create a new sheet and note the Spreadsheet ID from its URL: `https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID>/edit`.
   - Ensure it has a sheet/tab named `Sheet1` (or update `SHEET_NAME` in config).

4. Configure the app:
   - Open `config.py` and set:
     - `SPREADSHEET_ID = "YOUR_SPREADSHEET_ID"`
     - Optionally change `SHEET_NAME` and `GMAIL_QUERY`.

# Gmail to Google Sheets (Python)

Reads unread Gmail messages and appends Sender, Subject, Date, and Body as new rows in a Google Sheet. After logging, emails are marked as read. State is stored locally to avoid duplicates across runs.

---

## 1) High‑Level Architecture


## 2) Step‑by‑Step Setup (Windows)
1. Enable APIs
   - Open https://console.cloud.google.com/apis/dashboard
   - Create/select a project.
   - Enable: Gmail API and Google Sheets API.

2. Create OAuth credentials (Installed App)
   - Go to APIs & Services → Credentials → Create Credentials → OAuth client ID.
   - Application type: Desktop app.
   - Download the JSON and save it as `credentials/credentials.json`.

3. Create or choose a Google Sheet
   - Copy its Spreadsheet ID from the URL: `https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID>/edit`.
   - Confirm your tab (sheet) name (e.g., `Sheet1`).

4. Configure the app
   - Option A (env vars):
     ```powershell
     $env:GTS_SPREADSHEET_ID = "<YOUR_SPREADSHEET_ID>"
     $env:GTS_SHEET_NAME = "Sheet1"
     $env:GTS_GMAIL_QUERY = "is:unread in:inbox"
     ```
   - Option B (file): edit `config.py` and set `SPREADSHEET_ID`, `SHEET_NAME`.

5. Install and run
   ```powershell
   cd gmail-to-sheets
   python -m venv .venv
   .venv\Scripts\activate
   pip install --upgrade pip
   pip install -r requirements.txt
   python -m src.main
   ```
   - First run opens a browser for consent; token saved to `credentials/token.json`.

Notes
- If `SPREADSHEET_ID` is missing/invalid, the app auto‑creates a spreadsheet titled "Gmail To Sheets Log" with your tab name and prints the new ID.
- On an empty sheet, the script writes a header row: `From, Subject, Date, Body`.

---

## 3) How It Works
### 3.1 OAuth Flow Used
- OAuth 2.0 Installed Application (Desktop) via `google-auth-oauthlib`.
- Browser prompts once; a refreshable token is stored in `credentials/token.json`.
- You can set `GTS_OAUTH_MODE=no-open` to avoid auto‑opening the browser.

### 3.2 Duplicate Prevention Logic
- The script persists processed Gmail message IDs in `credentials/state.json`.
- Before processing, it skips IDs already in the state list.
- State is written only after Sheets append and mark‑as‑read succeed.

### 3.3 State Persistence Method
- File: `credentials/state.json`
- Fields: `processed_ids` (bounded list) and `last_run` (ISO timestamp).
- Cap controlled by `MAX_PROCESSED_IDS` in `config.py`.

---

## 4) Project Structure

## How It Prevents Duplicates
- The script stores a list of processed Gmail message IDs in `credentials/state.json`.
- Before processing, it skips any message whose ID is already recorded.
- It also filters with `is:unread` to reduce re-processing risk.
- State updates only after successful append + mark-as-read, so partial failures don’t corrupt state.

## State Storage
- File: `credentials/state.json`
- Fields:
  - `processed_ids`: list of message IDs appended to Sheets.
  - `last_run`: ISO timestamp of the last successful run.
- The list of processed IDs is capped (`MAX_PROCESSED_IDS`) to keep the file small.

## 5) Configuration Reference
- `SPREADSHEET_ID`: Target Google Sheet ID (or leave empty to auto‑create).
- `SHEET_NAME`: Tab name to append rows (created if missing).
- `GMAIL_QUERY`: Gmail search query (e.g., `is:unread in:inbox newer_than:7d`).
- `LOG_EVERY`: Progress print frequency during parsing (default 100).
- `BODY_MAX_CHARS`: Per‑cell body limit (default 50000; Sheets max).

Environment variables
```powershell
$env:GTS_SPREADSHEET_ID = "<your_id_here>"
$env:GTS_SHEET_NAME = "Sheet1"
$env:GTS_GMAIL_QUERY = "is:unread in:inbox"
$env:GTS_LOG_EVERY = "100"
$env:GTS_BODY_MAX_CHARS = "50000"
$env:GTS_OAUTH_MODE = "no-open"   # optional: do not auto-open browser
```

---

## 6) Challenges Faced & Solutions
- HTML‑only and multipart emails
  - Prefer `text/plain`; else, convert `text/html` to text with BeautifulSoup.
- Gmail/Sheets consent and test‑user restrictions
  - Documented consent screen setup; supports no‑open OAuth mode to ensure the correct account is chosen.
- Sheets 50k char per‑cell limit
  - Truncate body with a clear `... [truncated]` suffix when needed.
- Unknown/invalid Sheet ID
  - Auto‑create a spreadsheet and print the new ID; ensure tab exists.

---

## 7) Limitations
- Attachments are not captured; only message text is logged.
- HTML→text conversion is best‑effort and can lose complex formatting.
- Bodies over 50k characters are truncated due to Sheets limits.
- The `processed_ids` list is capped; extremely old unread emails outside the cap could reappear if still unread.
- Subject to Gmail/Sheets API quotas and rate limits.

---

## 8) Troubleshooting
- "Please set SPREADSHEET_ID": Provide a real ID or allow auto‑create.
- OAuth 403 access_denied: Add your account as a Test user or publish consent to Production; choose the correct account.
- Sheets 404: Ensure the ID is valid and your account has edit access.
- Long runs: Reduce scope with `GTS_GMAIL_QUERY` (e.g., `newer_than:7d`).

---

## 9) Quick Demo
```powershell
# Use an existing sheet
$env:GTS_SPREADSHEET_ID = "<YOUR_SHEET_ID>"
$env:GTS_SHEET_NAME = "Sheet1"
.venv\Scripts\python.exe -m src.main

# Or let the app create one and print the new ID
$env:GTS_SPREADSHEET_ID = ""
$env:GTS_SHEET_NAME = "InboxLog"
.venv\Scripts\python.exe -m src.main
```