from __future__ import annotations

import re
from typing import Callable

EMAIL_RE = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", re.I)
HEADER_RE = re.compile(r"^(from|sent|date|to)\s*:\s*(.*)$", re.I)
MARKER_RE = re.compile(r"^-{2,}\s*Original Message\s*-{2,}|^Begin forwarded message", re.I | re.M)


def classify_email(
    email: dict,
    tower_recipients: list[str],
    ground_recipients: list[str],
    ai_extract: Callable[[str], list[str]] | None = None,
    tower_senders: list[str] | None = None,
    ground_senders: list[str] | None = None,
) -> dict:
    """Classify Tower or Ground from the message.

    An Outlook forward is classified from the inner To line. A direct Gmail
    message is classified from the sender From address. AI may only supply
    addresses that already appear in the body, and only when no To line parsed.
    """
    email_id = str(email.get("internetMessageId") or email.get("id") or email.get("emailId") or "")
    attachments = email.get("attachments") or []
    try:
        return _classify(
            email,
            email_id,
            attachments,
            tower_recipients,
            ground_recipients,
            ai_extract,
            tower_senders or [],
            ground_senders or [],
        )
    except Exception as exc:
        return _review(email_id, attachments, f"Malformed email body: {exc}")


def _classify(
    email,
    email_id,
    attachments,
    tower_recipients,
    ground_recipients,
    ai_extract,
    tower_senders,
    ground_senders,
) -> dict:
    tower = {_norm(item) for item in tower_recipients}
    ground = {_norm(item) for item in ground_recipients}
    senders_tower = {_norm(item) for item in tower_senders}
    senders_ground = {_norm(item) for item in ground_senders}
    overlap = sorted((tower & ground) | (senders_tower & senders_ground))
    if overlap:
        return _review(email_id, attachments, f"Address is configured on both lists: {', '.join(overlap)}")
    if not tower and not ground and not senders_tower and not senders_ground:
        return _review(email_id, attachments, "Tower and Ground recipient and sender lists are empty")

    text = body_text(email)
    if _is_forward(text):
        if not text.strip():
            return _review(email_id, attachments, "Email body is missing or empty")
        return _classify_forward(email, email_id, attachments, text, tower, ground, ai_extract)
    if not text.strip() and not _object_addresses(email):
        return _review(email_id, attachments, "Email body is missing or empty")
    return _classify_sender(email, email_id, attachments, text, senders_tower, senders_ground)


def _classify_forward(email, email_id, attachments, text, tower, ground, ai_extract) -> dict:
    blocks = parse_blocks(text)
    chosen, to_source = _classification_block(blocks)
    extracted = list(dict.fromkeys(chosen["to"])) if chosen else []
    fallback = next((block for block in reversed(blocks) if block["from"]), blocks[-1] if blocks else None)
    if chosen and chosen.get("from"):
        original_from = chosen["from"][0]
    else:
        original_from = ((fallback or {}).get("from") or [""])[0] if fallback else ""
    if chosen and chosen.get("date"):
        original_date = chosen["date"]
    elif fallback:
        original_date = fallback.get("date") or ""
    else:
        original_date = ""

    if not extracted and not any(block["to"] for block in blocks) and ai_extract is not None:
        proposed = [_norm(item) for item in ai_extract(text)]
        extracted = [item for item in proposed if item and item in text.lower()]
        extracted = list(dict.fromkeys(extracted))
        to_source = "ai"

    base = {
        "emailId": email_id,
        "originalFrom": original_from,
        "originalDate": original_date,
        "body": text,
        "attachments": attachments,
        "subject": email.get("subject") or "",
        "toSource": to_source,
    }
    if not extracted:
        return {
            **base,
            "category": "Review",
            "originalTo": "",
            "reason": "Original To line could not be parsed",
            "status": "review_required",
            "needsAiToExtraction": ai_extract is None,
        }

    tower_hits = [item for item in extracted if item in tower]
    ground_hits = [item for item in extracted if item in ground]
    if tower_hits and ground_hits:
        return {
            **base,
            "category": "Review",
            "originalTo": "; ".join(tower_hits + ground_hits),
            "reason": "Forwarded To addresses match both Tower and Ground lists",
            "status": "review_required",
        }
    if tower_hits or ground_hits:
        hits = tower_hits or ground_hits
        category = "Tower" if tower_hits else "Ground"
        return {
            **base,
            "category": category,
            "originalTo": "; ".join(hits),
            "reason": f"Matched {category} recipient {', '.join(hits)}",
            "status": "classified",
        }
    return {
        **base,
        "category": "Review",
        "originalTo": "; ".join(extracted),
        "reason": f"Unrecognized forwarded To address: {', '.join(extracted)}",
        "status": "review_required",
    }


def _classify_sender(email, email_id, attachments, text, tower_senders, ground_senders) -> dict:
    senders = _sender_addresses(email, text)
    base = {
        "emailId": email_id,
        "originalFrom": "; ".join(senders),
        "originalDate": _direct_date(email, text),
        "body": text,
        "attachments": attachments,
        "subject": email.get("subject") or "",
        "toSource": "from",
        "originalTo": _direct_to(email, text),
    }
    if not senders:
        return {
            **base,
            "category": "Review",
            "reason": "Sender From address could not be read",
            "status": "review_required",
        }
    tower_hits = [item for item in senders if item in tower_senders]
    ground_hits = [item for item in senders if item in ground_senders]
    if tower_hits and ground_hits:
        return {
            **base,
            "category": "Review",
            "reason": "Sender addresses match both Tower and Ground lists",
            "status": "review_required",
        }
    if tower_hits or ground_hits:
        hits = tower_hits or ground_hits
        category = "Tower" if tower_hits else "Ground"
        return {
            **base,
            "originalFrom": "; ".join(hits),
            "category": category,
            "reason": f"Matched {category} sender {', '.join(hits)}",
            "status": "classified",
        }
    return {
        **base,
        "category": "Review",
        "reason": f"Unrecognized sender: {', '.join(senders)}",
        "status": "review_required",
    }


def _is_forward(text: str) -> bool:
    return MARKER_RE.search(text or "") is not None


def _sender_addresses(email: dict, text: str) -> list[str]:
    found = _object_addresses(email)
    if found:
        return found
    for key, value in _scan_headers(text):
        if key == "from":
            return _emails(value)
    return []


def _object_addresses(email: dict) -> list[str]:
    for key in ("from", "sender", "fromAddress"):
        found = _addresses_from_value(email.get(key))
        if found:
            return found
    return []


def _direct_to(email: dict, text: str) -> str:
    for key in ("to", "toAddress"):
        found = _addresses_from_value(email.get(key))
        if found:
            return "; ".join(found)
    for key, value in _scan_headers(text):
        if key == "to":
            found = _emails(value)
            if found:
                return "; ".join(found)
    return ""


def _direct_date(email: dict, text: str) -> str:
    for key in ("date", "originalDate"):
        value = email.get(key)
        if isinstance(value, str) and value.strip():
            return value.strip()
    for key, value in _scan_headers(text):
        if key in {"sent", "date"} and value:
            return value
    return ""


def _addresses_from_value(value) -> list[str]:
    if value is None or value == "":
        return []
    if isinstance(value, str):
        return _emails(value)
    if isinstance(value, dict):
        nested = value.get("emailAddress")
        if isinstance(nested, dict):
            return _addresses_from_value(nested.get("address") or nested.get("email") or "")
        for key in ("address", "email", "value"):
            if value.get(key):
                return _addresses_from_value(value[key])
        return []
    if isinstance(value, list):
        found: list[str] = []
        for item in value:
            found.extend(_addresses_from_value(item))
        return list(dict.fromkeys(found))
    return []


def _scan_headers(text: str) -> list[tuple[str, str]]:
    lines = (text or "").splitlines()
    headers: list[tuple[str, str]] = []
    index = 0
    while index < len(lines):
        line = lines[index]
        while index + 1 < len(lines) and re.match(r"^[ \t]", lines[index + 1]):
            nxt = lines[index + 1].strip()
            if re.match(r"^(from|sent|date|to|cc|bcc|subject)\s*:", nxt, re.I):
                break
            index += 1
            line += " " + nxt
        match = HEADER_RE.match(line.strip())
        index += 1
        if match:
            headers.append((match.group(1).lower(), match.group(2).strip()))
    return headers


def body_text(email: dict) -> str:
    body = email.get("body")
    if isinstance(body, str):
        return body
    if isinstance(body, dict) and isinstance(body.get("content"), str):
        content_type = str(body.get("contentType") or "").lower()
        content = body["content"]
        return html_to_text(content) if "html" in content_type else content
    for key in ("text", "textPlain", "bodyPreview"):
        if isinstance(email.get(key), str):
            return email[key]
    return ""


def html_to_text(html: str) -> str:
    text = re.sub(r"<br\s*/?>", "\n", html, flags=re.I)
    text = re.sub(r"</(p|div|tr|h\d|li|table)>", "\n", text, flags=re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    for src, dest in (
        ("&nbsp;", " "),
        ("&amp;", "&"),
        ("&lt;", "<"),
        ("&gt;", ">"),
        ("&#39;", "'"),
        ("&apos;", "'"),
        ("&quot;", '"'),
    ):
        text = text.replace(src, dest)
    return text


def _classification_block(blocks: list[dict]) -> tuple[dict | None, str]:
    """Use the original forwarded To line, not an intermediate hop."""
    if not blocks:
        return None, ""
    innermost_index = len(blocks) - 1
    for index, block in enumerate(blocks):
        if block["from"]:
            innermost_index = index
    if blocks[innermost_index]["to"]:
        return blocks[innermost_index], "innermost"
    steps = 1
    for index in range(innermost_index - 1, -1, -1):
        if blocks[index]["to"]:
            return blocks[index], f"outer_block_{steps}"
        steps += 1
    return None, ""


def parse_blocks(text: str) -> list[dict]:
    marker = MARKER_RE.search(text)
    region = text[marker.start() :] if marker else text
    lines = region.splitlines()
    blocks: list[dict] = []
    current = None

    def push() -> None:
        nonlocal current
        if current and (current["to"] or (current["from"] and current["date"])):
            blocks.append(current)

    index = 0
    while index < len(lines):
        line = lines[index]
        while index + 1 < len(lines) and re.match(r"^[ \t]", lines[index + 1]):
            nxt = lines[index + 1].strip()
            if re.match(r"^(from|sent|date|to|cc|bcc|subject)\s*:", nxt, re.I):
                break
            index += 1
            line += " " + nxt
        match = HEADER_RE.match(line.strip())
        index += 1
        if not match:
            continue
        key = match.group(1).lower()
        value = match.group(2).strip()
        if key == "from":
            push()
            current = {"from": _emails(value), "to": [], "date": ""}
        elif current is None:
            continue
        elif key == "to":
            current["to"].extend(_emails(value))
        else:
            current["date"] = value
    push()
    return blocks


def _emails(value: str) -> list[str]:
    return [_norm(match.group(0)) for match in EMAIL_RE.finditer(value or "")]


def _norm(address: str) -> str:
    return str(address or "").strip().lower()


def _review(email_id: str, attachments, reason: str) -> dict:
    return {
        "emailId": email_id,
        "category": "Review",
        "originalTo": "",
        "originalFrom": "",
        "originalDate": "",
        "reason": reason,
        "status": "review_required",
        "body": "",
        "attachments": attachments,
    }
