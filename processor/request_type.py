from __future__ import annotations

import json
import re
import urllib.request

from processor.settings import Settings
from processor.sow_catalog import SowCatalog
from processor.sow_model import SowClassifier


def classify_request_type(
    settings: Settings,
    category: str,
    sow: str,
    mop_text: str,
    email_body: str,
) -> dict:
    """Return a Request Type only when one allowed type is clear.

    Pass 1 uses the spreadsheet SOW. Pass 2 adds the MOP scope and the email
    body with addresses removed. Below the configured confidence, the type is blank.
    """
    allowed = list(settings.request_types.get(category) or [])
    phrases = settings.explicit_phrases.get(category) or []
    minimum = settings.request_type_min_confidence
    catalog = SowCatalog.load(settings.path(settings.catalog_path))
    model = SowClassifier.load(settings.path(settings.catalog_path))

    catalog_sow = _from_catalog(catalog, category, sow, allowed, minimum, "catalog_sow")
    if catalog_sow:
        return catalog_sow
    model_sow = _from_model(model, category, sow, allowed, minimum, "model_sow")
    if model_sow:
        return model_sow

    first = _explicit(sow, allowed, phrases)
    if first["confidence"] >= minimum:
        first["stage"] = "explicit_sow_only"
        return first

    if settings.ai_api_key:
        ai_first = _ai_classify(settings, category, allowed, sow, "", "")
        if ai_first["confidence"] >= minimum and ai_first["requestType"] in allowed:
            ai_first["stage"] = "sow_only"
            return ai_first

    combined_mop = _for_model(mop_text)
    combined_body = _for_model(email_body)
    catalog_mop = _from_catalog(catalog, category, combined_mop, allowed, minimum, "catalog_mop")
    if catalog_mop:
        return catalog_mop
    model_mop = _from_model(model, category, combined_mop, allowed, minimum, "model_mop")
    if model_mop:
        return model_mop

    second = _explicit(f"{sow}\n{combined_mop}\n{combined_body}", allowed, phrases)
    if second["confidence"] >= minimum:
        second["stage"] = "explicit_sow_mop_body"
        return second

    if settings.ai_api_key:
        ai_second = _ai_classify(settings, category, allowed, sow, combined_mop, combined_body)
        if ai_second["confidence"] >= minimum and ai_second["requestType"] in allowed:
            ai_second["stage"] = "sow_mop_body"
            return ai_second
        return {
            "requestType": "",
            "confidence": ai_second.get("confidence", 0),
            "stage": "unresolved",
            "reason": ai_second.get("reason") or f"Second pass confidence is below {minimum}",
        }

    return {
        "requestType": "",
        "confidence": max(first["confidence"], second["confidence"]),
        "stage": "unresolved",
        "reason": "No allowed Request Type was clear enough to assign",
    }


def _from_model(model: SowClassifier, category: str, text: str, allowed: list[str], minimum: int, stage: str) -> dict | None:
    hit = model.predict(category, text, allowed)
    if not hit or hit["confidence"] < minimum or not hit.get("requestType"):
        return None
    hit["stage"] = stage
    return hit


def _from_catalog(catalog: SowCatalog, category: str, text: str, allowed: list[str], minimum: int, stage: str) -> dict | None:
    hit = catalog.match(category, text, allowed)
    if not hit or hit["confidence"] < minimum:
        return None
    hit["stage"] = stage
    return hit


def prepare_brief_sow(sow: str, limit: int, settings: Settings) -> tuple[str, dict | None]:
    words = re.findall(r"\S+", sow or "")
    if len(words) <= limit:
        return sow or "", None
    if not settings.ai_api_key:
        return "", {
            "level": "site",
            "type": "sow_summary_unavailable",
            "message": f"BRIEF SOW is {len(words)} words and no AI summarizer is configured",
            "blocking": True,
        }
    summary = _summarize(settings, sow, limit)
    summary_words = re.findall(r"\S+", summary)
    if not summary or len(summary_words) > limit:
        return "", {
            "level": "site",
            "type": "sow_summary_failed",
            "message": f"SOW summary was empty or still over the {limit} word limit",
            "blocking": True,
        }
    return summary, {
        "level": "site",
        "type": "sow_summarized",
        "message": f"BRIEF SOW summarized from {len(words)} words to {len(summary_words)}",
        "blocking": False,
    }


def _explicit(text: str, allowed: list[str], phrases: list) -> dict:
    if not text or not text.strip():
        return {"requestType": "", "confidence": 0, "stage": "unresolved", "reason": "SOW is blank"}
    found: dict[str, str] = {}
    for entry in phrases:
        if not isinstance(entry, list) or len(entry) != 2:
            continue
        label, options = entry
        if label not in allowed:
            continue
        for phrase in options:
            if _contains(text, str(phrase)):
                found[label] = str(phrase)
                break
    if len(found) == 1:
        label, phrase = next(iter(found.items()))
        return {
            "requestType": label,
            "confidence": 95,
            "stage": "explicit",
            "reason": f"Matched phrase '{phrase}'",
        }
    if len(found) > 1:
        names = ", ".join(found)
        return {
            "requestType": "",
            "confidence": 40,
            "stage": "unresolved",
            "reason": f"SOW matches more than one allowed type: {names}",
        }
    return {"requestType": "", "confidence": 0, "stage": "unresolved", "reason": "No allowed phrase matched"}


def _contains(text: str, phrase: str) -> bool:
    if any(mark in phrase for mark in (" ", "/")):
        return phrase.casefold() in text.casefold()
    return re.search(rf"\b{re.escape(phrase)}\b", text, re.I) is not None


def _for_model(text: str) -> str:
    cleaned = re.sub(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", "", text or "", flags=re.I)
    cleaned = re.sub(r"\b\d{3}[-.)]\s*\d{3}[-.)]\s*\d{4}\b", "", cleaned)
    return cleaned.strip()


def _ai_classify(settings: Settings, category: str, allowed: list[str], sow: str, mop_text: str, body: str) -> dict:
    payload = {
        "category": category,
        "allowedTypes": allowed,
        "sow": sow,
        "mopSow": mop_text,
        "emailContext": body,
    }
    instruction = (
        "Classify the maintenance request. Return JSON with keys requestType and confidence. "
        "confidence is an integer from 0 to 100. requestType must be copied from allowedTypes "
        "or be an empty string. Do not invent a type. Do not return N/A."
    )
    try:
        raw = _chat(settings, instruction, json.dumps(payload))
        parsed = _json_object(raw)
    except Exception as exc:
        return {"requestType": "", "confidence": 0, "reason": f"AI classification failed: {exc}"}
    label = str(parsed.get("requestType") or "").strip()
    if label.lower() in {"n/a", "na", "none", "unknown"}:
        label = ""
    try:
        confidence = int(parsed.get("confidence") or 0)
    except (TypeError, ValueError):
        confidence = 0
    if label not in allowed:
        label = ""
        confidence = 0
    return {"requestType": label, "confidence": max(0, min(100, confidence)), "reason": "AI classification"}


def _summarize(settings: Settings, sow: str, limit: int) -> str:
    instruction = (
        f"Summarize the scope of work in {limit} words or fewer. Preserve the meaning and "
        "make the smallest change that fits the limit. Do not add tasks, dates, or names "
        "that are not in the source. Return JSON with key summary."
    )
    raw = _chat(settings, instruction, sow)
    parsed = _json_object(raw)
    return str(parsed.get("summary") or "").strip()


def _chat(settings: Settings, instruction: str, user_text: str) -> str:
    base = (settings.ai_base_url or "https://api.openai.com/v1").rstrip("/")
    body = json.dumps(
        {
            "model": settings.ai_model,
            "temperature": 0,
            "response_format": {"type": "json_object"},
            "messages": [
                {"role": "system", "content": instruction},
                {"role": "user", "content": user_text},
            ],
        }
    ).encode("utf-8")
    request = urllib.request.Request(
        f"{base}/chat/completions",
        data=body,
        headers={"Authorization": f"Bearer {settings.ai_api_key}", "Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=60) as response:
        payload = json.loads(response.read().decode("utf-8"))
    return payload["choices"][0]["message"]["content"]


def _json_object(raw: str) -> dict:
    try:
        value = json.loads(raw)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", raw, re.S)
        if not match:
            raise
        value = json.loads(match.group(0))
    if not isinstance(value, dict):
        raise ValueError("AI response was not a JSON object")
    return value

