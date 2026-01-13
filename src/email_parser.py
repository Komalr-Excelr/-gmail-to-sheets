import base64
import quopri
import re
from typing import Dict, List, Optional
from email.header import decode_header
from bs4 import BeautifulSoup


def _decode_base64url(data: Optional[str]) -> str:
    if not data:
        return ""
    # Gmail uses base64url (replace - and _) and may omit padding
    data = data.replace("-", "+").replace("_", "/")
    padding = 4 - (len(data) % 4)
    if padding and padding < 4:
        data += "=" * padding
    try:
        return base64.b64decode(data).decode("utf-8", errors="replace")
    except Exception:
        try:
            # Fallback quoted-printable decode if needed
            return quopri.decodestring(data).decode("utf-8", errors="replace")
        except Exception:
            return ""


def _decode_mime_words(s: str) -> str:
    parts = decode_header(s)
    decoded = []
    for text, enc in parts:
        if isinstance(text, bytes):
            try:
                decoded.append(text.decode(enc or "utf-8", errors="replace"))
            except Exception:
                decoded.append(text.decode("utf-8", errors="replace"))
        else:
            decoded.append(text)
    return "".join(decoded)


def extract_headers(headers: List[Dict[str, str]]) -> Dict[str, str]:
    wanted = {"from": "From", "subject": "Subject", "date": "Date"}
    result: Dict[str, str] = {}
    for h in headers:
        name = h.get("name", "")
        value = h.get("value", "")
        lname = name.lower()
        if lname in wanted:
            result[wanted[lname]] = _decode_mime_words(value)
    return result


def _walk_parts(payload: Dict) -> List[Dict]:
    parts = []
    if not payload:
        return parts
    mime_type = payload.get("mimeType")
    body = payload.get("body", {})
    data = body.get("data")
    parts.append({"mimeType": mime_type, "data": data})
    for p in payload.get("parts", []) or []:
        parts.extend(_walk_parts(p))
    return parts


def _html_to_text(html: str) -> str:
    soup = BeautifulSoup(html, "html.parser")
    # Reduce extra whitespace
    text = soup.get_text("\n")
    text = re.sub(r"\n\s*\n+", "\n\n", text)
    return text.strip()


def extract_body_text(payload: Dict) -> str:
    """Prefer text/plain; fallback to text/html converted to text."""
    parts = _walk_parts(payload)
    # Prefer text/plain
    for p in parts:
        if p.get("mimeType", "").startswith("text/plain"):
            return _decode_base64url(p.get("data"))
    # Fallback to HTML
    for p in parts:
        if p.get("mimeType", "").startswith("text/html"):
            html = _decode_base64url(p.get("data"))
            return _html_to_text(html)
    # If no parts, maybe body data is on root
    root_data = payload.get("body", {}).get("data")
    if root_data:
        return _decode_base64url(root_data)
    return ""
