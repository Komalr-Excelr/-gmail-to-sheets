from typing import Dict, List, Optional, Tuple
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from config import get_credentials, GMAIL_QUERY
from .email_parser import extract_headers, extract_body_text


def get_gmail_service():
    creds = get_credentials()
    # cache_discovery=False avoids writing cache files sometimes problematic on Windows
    return build("gmail", "v1", credentials=creds, cache_discovery=False)


def list_unread_message_ids(service, query: Optional[str] = None, max_results: int = 500) -> List[str]:
    """Return a list of unread message IDs in the inbox based on query."""
    query = query or GMAIL_QUERY
    user_id = "me"
    ids: List[str] = []
    page_token = None
    while True:
        try:
            req = service.users().messages().list(
                userId=user_id,
                q=query,
                maxResults=min(100, max_results),
                pageToken=page_token,
            )
            res = req.execute()
        except HttpError as e:
            raise RuntimeError(f"Gmail API error listing messages: {e}")
        messages = res.get("messages", [])
        ids.extend([m["id"] for m in messages])
        page_token = res.get("nextPageToken")
        if not page_token or len(ids) >= max_results:
            break
    return ids[:max_results]


def get_message_details(service, message_id: str) -> Dict[str, str]:
    """Fetch and parse message details for a given message ID."""
    try:
        msg = (
            service.users()
            .messages()
            .get(userId="me", id=message_id, format="full")
            .execute()
        )
    except HttpError as e:
        raise RuntimeError(f"Gmail API error getting message {message_id}: {e}")

    headers = extract_headers(msg.get("payload", {}).get("headers", []))
    sender = headers.get("From", "")
    subject = headers.get("Subject", "")
    date = headers.get("Date", "")
    body = extract_body_text(msg.get("payload", {}))
    internal_date_ms = msg.get("internalDate")  # str milliseconds

    return {
        "id": message_id,
        "from": sender,
        "subject": subject,
        "date": date,
        "body": body,
        "internalDate": internal_date_ms or "",
    }


def mark_messages_read(service, message_ids: List[str]) -> None:
    if not message_ids:
        return
    try:
        service.users().messages().batchModify(
            userId="me",
            body={"ids": message_ids, "removeLabelIds": ["UNREAD"]},
        ).execute()
    except HttpError as e:
        raise RuntimeError(f"Failed to mark messages as read: {e}")
