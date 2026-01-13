import json
import os
from datetime import datetime, timezone
from typing import Any, Dict, List

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from oauthlib.oauth2.rfc6749.errors import AccessDeniedError

# Scopes: Gmail modify (read + mark-as-read) and Sheets read/write
SCOPES: List[str] = [
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/spreadsheets",
]

# Spreadsheet configuration (update in README/setup). Can be overridden by env vars.
SPREADSHEET_ID = os.getenv("GTS_SPREADSHEET_ID", "123abcTEST")  # e.g., 1AbC...xyz
SHEET_NAME = os.getenv("GTS_SHEET_NAME", "Sheet1")  # change to your tab name
LOG_EVERY = int(os.getenv("GTS_LOG_EVERY", "100"))  # progress log frequency during parsing
BODY_MAX_CHARS = int(os.getenv("GTS_BODY_MAX_CHARS", "50000"))  # Sheets per-cell char limit

# Gmail search query: unread emails in Inbox (can override via GTS_GMAIL_QUERY)
GMAIL_QUERY = os.getenv("GTS_GMAIL_QUERY", "is:unread in:inbox")

# Files for OAuth tokens and state
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_DIR = os.path.join(BASE_DIR, "credentials")
CREDENTIALS_FILE = os.path.join(CREDENTIALS_DIR, "credentials.json")
TOKEN_FILE = os.path.join(CREDENTIALS_DIR, "token.json")
STATE_FILE = os.path.join(CREDENTIALS_DIR, "state.json")

# State retention
MAX_PROCESSED_IDS = 5000


def ensure_credentials_dir() -> None:
    os.makedirs(CREDENTIALS_DIR, exist_ok=True)


def get_credentials() -> Credentials:
    """Load stored credentials or run local OAuth flow to obtain new ones."""
    ensure_credentials_dir()
    creds = None
    if os.path.exists(TOKEN_FILE):
        try:
            creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
        except Exception:
            creds = None
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                raise FileNotFoundError(
                    f"Missing OAuth client secrets at {CREDENTIALS_FILE}. See README to create it."
                )
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            oauth_mode = os.getenv("GTS_OAUTH_MODE", "local").lower()
            try:
                if oauth_mode == "no-open":
                    # Print URL but do not auto-open the browser
                    creds = flow.run_local_server(port=0, open_browser=False)
                else:
                    # Opens a browser and listens on localhost
                    creds = flow.run_local_server(port=0)
            except AccessDeniedError:
                print("OAuth was denied. Ensure your account is added as a Test user or publish the app, then retry.")
                raise
        # Save the credentials for next run
        with open(TOKEN_FILE, "w", encoding="utf-8") as token:
            token.write(creds.to_json())
    return creds


def load_state() -> Dict[str, Any]:
    ensure_credentials_dir()
    if not os.path.exists(STATE_FILE):
        return {"processed_ids": [], "last_run": None}
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"processed_ids": [], "last_run": None}


def save_state(state: Dict[str, Any]) -> None:
    # trim processed IDs list to MAX_PROCESSED_IDS
    if isinstance(state.get("processed_ids"), list) and len(state["processed_ids"]) > MAX_PROCESSED_IDS:
        state["processed_ids"] = state["processed_ids"][-MAX_PROCESSED_IDS:]
    state["last_run"] = datetime.now(timezone.utc).isoformat()
    ensure_credentials_dir()
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)
