import sys
from typing import List

from config import (
    load_state,
    save_state,
    SPREADSHEET_ID,
    SHEET_NAME,
    GMAIL_QUERY,
    LOG_EVERY,
    BODY_MAX_CHARS,
)
from .gmail_service import get_gmail_service, list_unread_message_ids, get_message_details, mark_messages_read
from .sheets_service import (
    get_sheets_service,
    append_rows,
    ensure_header_row,
    ensure_spreadsheet_and_sheet,
)


def main() -> int:
    if SPREADSHEET_ID == "YOUR_SPREADSHEET_ID":
        print("ERROR: Please set SPREADSHEET_ID in config.py before running.")
        return 2

    state = load_state()
    processed_ids: List[str] = state.get("processed_ids", [])
    processed_set = set(processed_ids)

    gmail = get_gmail_service()
    sheets = get_sheets_service()

    print("Fetching unread messages from Gmail...")
    message_ids = list_unread_message_ids(gmail, GMAIL_QUERY, max_results=1000)
    print(f"Found {len(message_ids)} unread messages.")

    new_rows = []
    to_mark_read: List[str] = []

    total = len(message_ids)
    for idx, mid in enumerate(message_ids, start=1):
        if mid in processed_set:
            # already processed previously; skip to prevent duplicates
            continue
        details = get_message_details(gmail, mid)
        body = details.get("body", "").replace("\r\n", "\n")
        # Truncate body to avoid Sheets 50k char per-cell limit
        if len(body) > BODY_MAX_CHARS:
            suffix = "... [truncated]"
            body = body[: max(0, BODY_MAX_CHARS - len(suffix))] + suffix

        row = [
            details.get("from", ""),
            details.get("subject", ""),
            details.get("date", ""),
            body,
        ]
        new_rows.append(row)
        to_mark_read.append(mid)
        if idx % max(1, LOG_EVERY) == 0:
            print(f"Parsed {idx}/{total} messages...")

    if not new_rows:
        print("No new unread messages to append. Exiting.")
        return 0

    # Ensure spreadsheet and sheet exist, creating them if needed
    target_spreadsheet_id = ensure_spreadsheet_and_sheet(sheets, SPREADSHEET_ID, SHEET_NAME)

    # Ensure header row if sheet is empty
    ensure_header_row(sheets, ["From", "Subject", "Date", "Body"], spreadsheet_id=target_spreadsheet_id, sheet_name=SHEET_NAME)

    print(f"Appending {len(new_rows)} rows to Google Sheets...")
    append_rows(sheets, new_rows, spreadsheet_id=target_spreadsheet_id, sheet_name=SHEET_NAME)

    print(f"Marking {len(to_mark_read)} emails as read in Gmail...")
    mark_messages_read(gmail, to_mark_read)

    # update state only after successful append and mark-as-read
    processed_ids.extend(to_mark_read)
    state["processed_ids"] = processed_ids
    save_state(state)

    print("Done.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
