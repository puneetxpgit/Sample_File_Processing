from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path

from processor.export_kirke import (
    iso_date,
    iso_time,
    rebuild_category,
    reset_category,
    safe_id,
    schedule_key,
    write_manifest,
    write_reply,
)
from processor.match_mops import filename_digit_runs, match_rows
from processor.read_inputs import read_mop, read_spreadsheet
from processor.request_type import classify_request_type, prepare_brief_sow
from processor.settings import Settings

EMAIL_RE = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", re.I)


def process_package(settings: Settings, job: dict) -> dict:
    email_id = str(job.get("emailId") or "").strip() or "email"
    category = str(job.get("category") or "Review").strip()
    if category not in {"Tower", "Ground", "Review"}:
        category = "Review"
    original_from = _email_only(job.get("originalFrom") or "")
    original_to = str(job.get("originalTo") or "")
    original_date_raw = str(job.get("originalDate") or "")
    body = str(job.get("body") or "")
    request_date = _parse_request_date(original_date_raw) or _date_from_body(body)
    spreadsheet_paths = [Path(item) for item in _as_list(job.get("spreadsheet") or job.get("spreadsheets"))]
    mop_paths = [Path(item) for item in _as_list(job.get("mops") or job.get("mop"))]

    issues: list[dict] = []
    if category in {"Tower", "Ground"} and not request_date:
        issues.append(_issue("email", "original_date_missing", "Original forwarded Date header is missing. Request Date was left blank.", blocking=False))
    if category in {"Tower", "Ground"} and not original_from:
        issues.append(_issue("email", "original_from_missing", "Original forwarded From address is missing. Requested by GC Email was left blank.", blocking=False))

    if not spreadsheet_paths:
        issues.append(_issue("email", "no_spreadsheet", "No spreadsheet attachment was found. This email was skipped."))
        result = _result(job, category, original_from, original_to, original_date_raw, "failed", issues, [], {})
        _remember(settings, result)
        return result

    rows = []
    invalid_rows = 0
    try:
        for path in spreadsheet_paths:
            parsed, warnings = read_spreadsheet(path, settings.header_synonyms)
            rows.extend(parsed)
            for warning in warnings:
                invalid_rows += 1
                issues.append(_issue("site", "invalid_row", warning))
    except Exception as exc:
        issues.append(_issue("email", "spreadsheet_failed", f"Spreadsheet could not be read: {exc}"))
        result = _result(job, category, original_from, original_to, original_date_raw, "failed", issues, [], {})
        _remember(settings, result)
        return result

    documents = []
    failed_paths: list[tuple[Path, str]] = []
    for path in mop_paths:
        try:
            documents.append(read_mop(path))
        except Exception as exc:
            failed_paths.append((path, str(exc)))
            issues.append(_issue("attachment", "read_failed", str(exc)))

    pairs, unmatched = match_rows(rows, documents, settings.match_strategy)
    output_rows = []
    included = 0
    excluded = 0
    for sheet_row, mop, reason in pairs:
        ambiguous_names = ""
        if reason.startswith("ambiguous_mop|"):
            ambiguous_names = reason.split("|", 1)[1]
            reason = "ambiguous_mop"
        if reason == "missing_mop":
            failed = _failed_match(sheet_row, failed_paths)
            if failed is not None:
                reason = "read_failed"
                message = f"MOP {failed.name} failed to read for Fuze Project ID {sheet_row.fuze_id}"
            else:
                message = f"No MOP matched Fuze Project ID {sheet_row.fuze_id}"
        elif reason == "ambiguous_mop":
            message = f"More than one MOP matched Fuze Project ID {sheet_row.fuze_id}: {ambiguous_names}"
        else:
            message = ""

        if mop is None:
            excluded += 1
            issues.append(_issue("site", reason, message, sheet_row.fuze_id))
            continue

        for warning in sheet_row.warnings:
            issues.append(_issue("site", "field_warning", warning, sheet_row.fuze_id, blocking=False))

        brief, sow_issue = prepare_brief_sow(sheet_row.sow, settings.brief_sow_word_limit, settings)
        if sow_issue and sow_issue.get("blocking"):
            excluded += 1
            sow_issue["fuzeProjectId"] = sheet_row.fuze_id
            issues.append(sow_issue)
            continue
        if sow_issue:
            sow_issue["fuzeProjectId"] = sheet_row.fuze_id
            issues.append(sow_issue)

        if category in {"Tower", "Ground"}:
            classified = classify_request_type(settings, category, sheet_row.sow, mop.scope_text, body)
        else:
            classified = {
                "requestType": "",
                "confidence": 0,
                "stage": "review",
                "reason": "Category is Review, so Request Type was left blank",
            }
        included += 1
        output_rows.append(
            {
                "emailId": email_id,
                "sourceRow": sheet_row.source_row,
                "requestDate": request_date,
                "startDate": iso_date(sheet_row.start_date),
                "endDate": iso_date(sheet_row.end_date),
                "startTime": iso_time(sheet_row.start_time),
                "endTime": iso_time(sheet_row.end_time),
                "fuzeProjectId": sheet_row.fuze_id,
                "psLocation": sheet_row.ps_location,
                "enodeb": sheet_row.enodeb,
                "siteName": sheet_row.site_name,
                "gcVendor": sheet_row.gc_vendor,
                "gcEmail": original_from,
                "briefSow": brief,
                "requestType": classified["requestType"],
                "requestTypeConfidence": classified["confidence"],
                "requestTypeStage": classified["stage"],
                "requestTypeReason": classified["reason"],
                "matchReason": reason,
                "mopSource": str(mop.path),
            }
        )

    for mop in unmatched:
        issues.append(
            _issue(
                "attachment",
                "unmatched_mop",
                f"{mop.path.name} has no matching spreadsheet row",
                blocking=False,
            )
        )

    category_rows, review_rows = _split_untyped(category, output_rows, issues)
    rows_added, rows_updated = _count_changes(settings, category, email_id, category_rows)
    review_added, review_updated = _count_changes(settings, "Review", email_id, review_rows)
    rows_added += review_added
    rows_updated += review_updated
    for row in _updated_rows(settings, category, email_id, category_rows) + _updated_rows(
        settings, "Review", email_id, review_rows
    ):
        stamp = _short_date(row.get("startDate") or "")
        issues.append(
            _issue(
                "site",
                "duplicate_row",
                f"Fuze {row.get('fuzeProjectId')} {stamp} already scheduled – updated existing row.",
                str(row.get("fuzeProjectId") or ""),
                blocking=False,
            )
        )

    if category == "Review" or (review_rows and not category_rows):
        email_status = "review_required"
        processing_status = "review_required"
    elif review_rows:
        email_status = "review_required"
        processing_status = "review_required"
    elif category in {"Tower", "Ground"}:
        email_status = "pending_delivery"
        processing_status = _processing_status(len(category_rows), excluded + invalid_rows)
    else:
        email_status = "review_required"
        processing_status = "review_required"

    outputs = _publish(
        settings,
        email_id,
        category,
        original_to,
        original_from,
        original_date_raw,
        category_rows,
        review_rows,
    )
    primary_rows = category_rows or review_rows
    result = _result(
        job,
        category,
        original_from,
        original_to,
        original_date_raw,
        processing_status,
        issues,
        primary_rows,
        outputs,
        email_status=email_status,
        included=included,
        excluded=excluded + invalid_rows,
        invalid_rows=invalid_rows,
        rows_added=rows_added,
        rows_updated=rows_updated,
        rows_skipped=excluded + invalid_rows,
        review_rows=review_rows if category_rows else [],
    )
    _remember(settings, result)
    return result


def run_samples(settings: Settings) -> list[dict]:
    for category in ("Tower", "Ground", "Review"):
        reset_category(settings, category)
    results = []
    for job in settings.sample_jobs:
        package = dict(job)
        package["spreadsheet"] = str(settings.path(job["spreadsheet"]))
        package["mops"] = [str(settings.path(item)) for item in job.get("mops") or []]
        package["body"] = job.get("body") or ""
        results.append(process_package(settings, package))
    report = inspect_workspace(settings)
    report_path = settings.path(settings.output_dir) / "inspect-report.json"
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, indent=2), encoding="utf-8")
    return results


def inspect_workspace(settings: Settings) -> dict:
    spreadsheets = []
    mops = []
    for path in sorted(settings.root.iterdir()):
        if not path.is_file():
            continue
        lower = path.name.casefold()
        if path.suffix.lower() in {".xlsx", ".xlsm", ".xls", ".csv"} and "output format" not in lower:
            try:
                rows, warnings = read_spreadsheet(path, settings.header_synonyms)
                spreadsheets.append(
                    {
                        "file": path.name,
                        "rows": [
                            {
                                "sourceRow": row.source_row,
                                "fuzeProjectId": row.fuze_id,
                                "siteName": row.site_name,
                                "gcVendor": row.gc_vendor,
                                "startDate": iso_date(row.start_date),
                                "sowWords": len(re.findall(r"\S+", row.sow)),
                            }
                            for row in rows
                        ],
                        "excluded": warnings,
                    }
                )
            except Exception as exc:
                spreadsheets.append({"file": path.name, "error": str(exc)})
        elif path.suffix.lower() == ".docx":
            try:
                mop = read_mop(path)
                mops.append(
                    {
                        "file": path.name,
                        "fuzeIds": mop.fuze_ids,
                        "siteName": mop.site_name,
                        "psLocation": mop.ps_location,
                        "jobDescription": mop.job_description,
                        "dateWindow": mop.date_window,
                        "hasScopeOfWork": mop.scope_text != mop.text,
                    }
                )
            except Exception as exc:
                mops.append({"file": path.name, "error": str(exc)})

    schedule_ids = {
        row["fuzeProjectId"]
        for book in spreadsheets
        for row in book.get("rows") or []
    }
    unpaired = []
    for mop in mops:
        ids = set(mop.get("fuzeIds") or [])
        if ids and ids.isdisjoint(schedule_ids):
            unpaired.append(mop["file"])
        elif not ids:
            site = (mop.get("siteName") or "").casefold()
            known_sites = {
                row["siteName"].casefold()
                for book in spreadsheets
                for row in book.get("rows") or []
                if row.get("siteName")
            }
            if site and site not in known_sites:
                unpaired.append(mop["file"])
    return {"spreadsheets": spreadsheets, "mops": mops, "mopsWithoutSpreadsheet": unpaired}


def _split_untyped(category: str, rows: list[dict], issues: list[dict]) -> tuple[list[dict], list[dict]]:
    if category not in {"Tower", "Ground"}:
        return rows, []
    kept = []
    review = []
    for row in rows:
        if str(row.get("requestType") or "").strip():
            kept.append(row)
            continue
        review.append(row)
        sow = str(row.get("briefSow") or "").strip() or "blank"
        if len(sow) > 140:
            sow = sow[:137] + "..."
        site = row.get("siteName") or ""
        issues.append(
            _issue(
                "site",
                "request_type_unmapped",
                (
                    f"No {category} request type matches SOW '{sow}' "
                    f"(Fuze {row.get('fuzeProjectId')}, {site}). Check the category."
                ),
                str(row.get("fuzeProjectId") or ""),
                blocking=False,
            )
        )
    return kept, review


def _publish(
    settings: Settings,
    email_id: str,
    category: str,
    original_to: str,
    original_from: str,
    original_date: str,
    category_rows: list[dict],
    review_rows: list[dict],
) -> dict:
    outputs: dict = {}
    if category == "Review":
        wrote = _store_manifest(
            settings, "Review", email_id, original_to, original_from, original_date, category_rows
        )
        if category_rows or wrote:
            master = _refresh(settings, "Review", email_id, category_rows)
            outputs["masterWorkbook"] = master.get("workbook", "")
            if category_rows:
                outputs["workbook"] = master.get("replyWorkbook", "")
                outputs["replyWorkbook"] = master.get("replyWorkbook", "")
                outputs["mops"] = master.get("replyMops") or []
                outputs["rowCount"] = master.get("rowCount", 0)
                outputs["mopCount"] = master.get("mopCount", 0)
        return outputs

    wrote_category = _store_manifest(
        settings, category, email_id, original_to, original_from, original_date, category_rows
    )
    if category_rows or wrote_category:
        master = _refresh(settings, category, email_id, category_rows)
        outputs["masterWorkbook"] = master.get("workbook", "")
        if category_rows:
            outputs["workbook"] = master.get("replyWorkbook", "")
            outputs["replyWorkbook"] = master.get("replyWorkbook", "")
            outputs["mops"] = master.get("replyMops") or []
            outputs["rowCount"] = master.get("rowCount", 0)
            outputs["mopCount"] = master.get("mopCount", 0)

    diverted = category in {"Tower", "Ground"}
    wrote_review = False
    if diverted:
        wrote_review = _store_manifest(
            settings, "Review", email_id, original_to, original_from, original_date, review_rows
        )
    if review_rows or wrote_review:
        review = _refresh(settings, "Review", email_id, review_rows)
        if review_rows and category_rows:
            outputs["reviewWorkbook"] = review.get("replyWorkbook", "")
            outputs["reviewReplyWorkbook"] = review.get("replyWorkbook", "")
            outputs["reviewMops"] = review.get("replyMops") or []
        elif review_rows:
            outputs["workbook"] = review.get("replyWorkbook", "")
            outputs["replyWorkbook"] = review.get("replyWorkbook", "")
            outputs["mops"] = review.get("replyMops") or []
            outputs["masterWorkbook"] = review.get("workbook", "")
            outputs["rowCount"] = review.get("rowCount", 0)
            outputs["mopCount"] = review.get("mopCount", 0)
    return outputs


def _store_manifest(
    settings: Settings,
    category: str,
    email_id: str,
    original_to: str,
    original_from: str,
    original_date: str,
    rows: list[dict],
) -> bool:
    """Write this email's manifest, or remove it when this run has no rows. Returns True if a file was removed."""
    path = settings.path(settings.output_dir) / category / "manifests" / f"{safe_id(email_id)}.json"
    if not rows:
        if path.exists():
            path.unlink()
            return True
        return False
    write_manifest(
        settings,
        category,
        email_id,
        {
            "emailId": email_id,
            "category": category,
            "originalTo": original_to,
            "originalFrom": original_from,
            "originalDate": original_date,
            "rows": rows,
        },
    )
    return False


def _refresh(settings: Settings, category: str, email_id: str, rows: list[dict]) -> dict:
    master = rebuild_category(settings, category)
    if not rows:
        return master
    reply = write_reply(settings, category, email_id, rows)
    master["replyWorkbook"] = reply["workbook"]
    master["replyMops"] = reply["mops"]
    return master


def _count_changes(settings: Settings, category: str, email_id: str, rows: list[dict]) -> tuple[int, int]:
    existing = _existing_keys(settings, category, email_id)
    added = 0
    updated = 0
    seen: set[tuple] = set()
    for row in rows:
        key = schedule_key(row)
        if key in seen:
            continue
        seen.add(key)
        if key in existing:
            updated += 1
        else:
            added += 1
    return added, updated


def _updated_rows(settings: Settings, category: str, email_id: str, rows: list[dict]) -> list[dict]:
    existing = _existing_keys(settings, category, email_id)
    found = []
    seen: set[tuple] = set()
    for row in rows:
        key = schedule_key(row)
        if key in seen or key not in existing:
            continue
        seen.add(key)
        found.append(row)
    return found


def _existing_keys(settings: Settings, category: str, email_id: str) -> set[tuple]:
    folder = settings.path(settings.output_dir) / category / "manifests"
    if not folder.exists():
        return set()
    own = f"{safe_id(email_id)}.json"
    keys: set[tuple] = set()
    for path in folder.glob("*.json"):
        if path.name == own:
            continue
        data = json.loads(path.read_text(encoding="utf-8"))
        for row in data.get("rows") or []:
            keys.add(schedule_key(row))
    return keys


def _short_date(iso_value: str) -> str:
    match = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})", iso_value or "")
    if not match:
        return iso_value or "undated"
    return f"{int(match.group(2))}/{int(match.group(3))}"


def _processing_status(included: int, excluded: int) -> str:
    if included == 0:
        return "incomplete"
    if excluded:
        return "partial"
    return "completed"


def _failed_match(row, failed_paths: list[tuple[Path, str]]) -> Path | None:
    site_tokens = [part for part in re.split(r"[^a-z0-9]+", row.site_name.casefold()) if len(part) >= 2]
    for path, _message in failed_paths:
        if row.fuze_id and row.fuze_id in filename_digit_runs(path.stem):
            return path
        name_tokens = [part for part in re.split(r"[^a-z0-9]+", path.stem.casefold()) if part]
        if site_tokens and all(token in name_tokens for token in site_tokens):
            return path
    return None


def _result(
    job: dict,
    category: str,
    original_from: str,
    original_to: str,
    original_date: str,
    processing_status: str,
    issues: list[dict],
    output_rows: list[dict],
    outputs: dict,
    email_status: str | None = None,
    included: int | None = None,
    excluded: int | None = None,
    invalid_rows: int = 0,
    rows_added: int = 0,
    rows_updated: int = 0,
    rows_skipped: int = 0,
    review_rows: list[dict] | None = None,
) -> dict:
    if email_status is None:
        email_status = "failed" if processing_status == "failed" else "pending_delivery"
    mop_names = sorted({row.get("mopFileName") for row in output_rows if row.get("mopFileName")})
    return {
        "emailId": job.get("emailId") or "",
        "category": category,
        "originalTo": original_to,
        "originalFrom": original_from,
        "originalDate": original_date,
        "reason": job.get("categoryNote") or job.get("reason") or "",
        "status": "review_required" if category == "Review" else "classified",
        "processingStatus": processing_status,
        "emailStatus": email_status,
        "rowCount": len(output_rows),
        "mopCount": len(mop_names),
        "includedSites": included if included is not None else len(output_rows),
        "excludedSites": excluded if excluded is not None else 0,
        "invalidRows": invalid_rows,
        "rowsAdded": rows_added,
        "rowsUpdated": rows_updated,
        "rowsSkipped": rows_skipped,
        "issues": [_with_code(issue) for issue in issues],
        "rows": [
            {
                "fuzeProjectId": row["fuzeProjectId"],
                "siteName": row["siteName"],
                "requestType": row["requestType"],
                "confidence": row["requestTypeConfidence"],
                "stage": row["requestTypeStage"],
                "reason": row["requestTypeReason"],
                "mopFileName": row.get("mopFileName", ""),
            }
            for row in output_rows
        ],
        "reviewRows": [
            {
                "fuzeProjectId": row["fuzeProjectId"],
                "siteName": row["siteName"],
                "requestType": row["requestType"],
                "mopFileName": row.get("mopFileName", ""),
            }
            for row in (review_rows or [])
        ],
        "outputs": outputs,
    }


def _with_code(issue: dict) -> dict:
    coded = dict(issue)
    if not coded.get("code"):
        coded["code"] = {
            "original_date_missing": "DATE_MISSING",
            "duplicate_row": "DUPLICATE_ROW",
            "request_type_unmapped": "REQUEST_TYPE_UNMAPPED",
        }.get(str(coded.get("type") or ""), str(coded.get("type") or "ISSUE").upper())
    return coded


def _remember(settings: Settings, result: dict) -> None:
    log_path = settings.path(settings.output_dir) / "outlook-automation-log.jsonl"
    log_path.parent.mkdir(parents=True, exist_ok=True)
    with log_path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(result) + "\n")


def _issue(level: str, issue_type: str, message: str, fuze_project_id: str = "", blocking: bool = True) -> dict:
    return {
        "level": level,
        "type": issue_type,
        "message": message,
        "fuzeProjectId": fuze_project_id,
        "blocking": blocking,
    }


def _as_list(value) -> list:
    if value is None or value == "":
        return []
    if isinstance(value, list):
        return value
    return [part.strip() for part in str(value).split("|") if part.strip()]


def _email_only(value: str) -> str:
    match = EMAIL_RE.search(value or "")
    return match.group(0) if match else ""


def _date_from_body(body: str) -> str:
    for line in (body or "").splitlines():
        match = re.match(r"^(?:date|sent)\s*:\s*(.*)$", line.strip(), re.I)
        if not match:
            continue
        parsed = _parse_request_date(match.group(1))
        if parsed:
            return parsed
    return ""


def _parse_request_date(value: str) -> str:
    text = (value or "").strip()
    text = re.sub(r"^(?:sent|date)\s*:\s*", "", text, flags=re.I).strip()
    if not text:
        return ""
    iso = re.fullmatch(
        r"(\d{4}-\d{2}-\d{2})(?:T\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:\d{2})?)?",
        text,
    )
    if iso:
        return _valid_date(iso.group(1), "%Y-%m-%d")
    slash = re.fullmatch(
        r"(\d{1,2}/\d{1,2}/\d{2}(?:\d{2})?)(?:\s+\d{1,2}:\d{2}\s*[AP]M)?",
        text,
        re.I,
    )
    if slash:
        numeric = slash.group(1)
        pattern = "%m/%d/%y" if len(numeric.rsplit("/", 1)[-1]) == 2 else "%m/%d/%Y"
        return _valid_date(numeric, pattern)
    month_first = re.fullmatch(
        r"(?:[A-Za-z]+,\s+)?([A-Za-z]+)\.?\s+(\d{1,2}),\s+(\d{4})(?:\s+at)?(?:\s+\d{1,2}:\d{2}(?::\d{2})?\s*[AP]M)?",
        text,
        re.I,
    )
    if month_first:
        month, day, year = month_first.group(1), month_first.group(2), month_first.group(3)
        return _month_name_date(month, day, year)
    day_first = re.fullmatch(
        r"(?:[A-Za-z]+,\s+)?(\d{1,2})\s+([A-Za-z]+)\.?\s+(\d{4})(?:\s+at)?(?:\s+\d{1,2}:\d{2}(?::\d{2})?\s*[AP]M)?",
        text,
        re.I,
    )
    if day_first:
        day, month, year = day_first.group(1), day_first.group(2), day_first.group(3)
        return _month_name_date(month, day, year)
    return ""


def _month_name_date(month: str, day: str, year: str) -> str:
    for pattern in ("%B", "%b"):
        found = _valid_date(f"{month} {day} {year}", f"{pattern} %d %Y")
        if found:
            return found
    return ""


def _valid_date(text: str, pattern: str) -> str:
    try:
        return datetime.strptime(text, pattern).date().isoformat()
    except ValueError:
        return ""

