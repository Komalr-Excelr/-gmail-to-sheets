from typing import List, Optional
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from config import get_credentials, SPREADSHEET_ID, SHEET_NAME


def get_sheets_service():
    creds = get_credentials()
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def append_rows(service, values: List[List[str]], spreadsheet_id: str = None, sheet_name: str = None) -> None:
    if not values:
        return
    spreadsheet_id = spreadsheet_id or SPREADSHEET_ID
    sheet_name = sheet_name or SHEET_NAME
    rng = f"{sheet_name}!A:D"
    try:
        service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=rng,
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": values},
        ).execute()
    except HttpError as e:
        raise RuntimeError(f"Sheets API append error: {e}")


def get_first_row(service, spreadsheet_id: Optional[str] = None, sheet_name: Optional[str] = None) -> List[str]:
    spreadsheet_id = spreadsheet_id or SPREADSHEET_ID
    sheet_name = sheet_name or SHEET_NAME
    rng = f"{sheet_name}!A1:D1"
    try:
        res = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=rng,
        ).execute()
        values = res.get("values", [])
        return values[0] if values else []
    except HttpError as e:
        raise RuntimeError(f"Sheets API read error: {e}")


def ensure_header_row(service, headers: List[str], spreadsheet_id: Optional[str] = None, sheet_name: Optional[str] = None) -> None:
    spreadsheet_id = spreadsheet_id or SPREADSHEET_ID
    sheet_name = sheet_name or SHEET_NAME
    first = get_first_row(service, spreadsheet_id, sheet_name)
    if first:
        return
    append_rows(service, [headers], spreadsheet_id, sheet_name)


def get_spreadsheet_metadata(service, spreadsheet_id: str) -> dict:
    try:
        return service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    except HttpError as e:
        raise RuntimeError(f"Sheets API metadata error: {e}")


def ensure_spreadsheet_and_sheet(service, spreadsheet_id: Optional[str], sheet_name: str) -> str:
    """Ensure spreadsheet ID exists and sheet/tab is present. Create as needed.

    Returns the valid spreadsheet ID to use for this run.
    """
    # If spreadsheet_id looks placeholder, force create
    need_create = not spreadsheet_id or spreadsheet_id == "YOUR_SPREADSHEET_ID" or spreadsheet_id == "123abcTEST"
    if not need_create:
        try:
            meta = get_spreadsheet_metadata(service, spreadsheet_id)
            titles = [s["properties"]["title"] for s in meta.get("sheets", [])]
            if sheet_name not in titles:
                # add the missing sheet
                try:
                    service.spreadsheets().batchUpdate(
                        spreadsheetId=spreadsheet_id,
                        body={
                            "requests": [
                                {"addSheet": {"properties": {"title": sheet_name}}}
                            ]
                        },
                    ).execute()
                except HttpError as e:
                    raise RuntimeError(f"Failed adding sheet '{sheet_name}': {e}")
            return spreadsheet_id
        except RuntimeError:
            need_create = True

    # Create new spreadsheet
    body = {
        "properties": {"title": "Gmail To Sheets Log"},
        "sheets": [{"properties": {"title": sheet_name}}],
    }
    try:
        res = service.spreadsheets().create(body=body, fields="spreadsheetId").execute()
        new_id = res.get("spreadsheetId")
        print(f"Created new spreadsheet with ID: {new_id}. Update config or set GTS_SPREADSHEET_ID to reuse.")
        return new_id
    except HttpError as e:
        raise RuntimeError(f"Failed to create spreadsheet: {e}")
