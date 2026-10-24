from __future__ import annotations

import argparse
import base64
import json
import sys
from pathlib import Path

from processor.classify_email import classify_email
from processor.pipeline import inspect_workspace, process_package, run_samples
from processor.settings import load_settings
from processor.sow_model import SowClassifier


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Outlook MOP automation")
    parser.add_argument("--root", type=Path, default=None)
    parser.add_argument("--config", type=Path, default=None)
    sub = parser.add_subparsers(dest="command", required=True)

    classify = sub.add_parser("classify-email", help="Classify a forwarded email body")
    classify.add_argument("--body-file", type=Path)
    classify.add_argument("--body-b64", default="")
    classify.add_argument("--email-id", default="")
    classify.add_argument("--from", dest="sender", default="")
    classify.add_argument("--tower", default="")
    classify.add_argument("--ground", default="")

    process = sub.add_parser("process", help="Process one classified email package")
    process.add_argument("--email-id", required=True)
    process.add_argument("--category", required=True, choices=["Tower", "Ground", "Review"])
    process.add_argument("--original-to", default="")
    process.add_argument("--original-from", default="")
    process.add_argument("--original-date", default="")
    process.add_argument("--reason", default="")
    process.add_argument("--body-file", type=Path)
    process.add_argument("--body-b64", default="")
    process.add_argument("--spreadsheet", action="append", default=[])
    process.add_argument("--mop", action="append", default=[])

    sow = sub.add_parser("classify-sow", help="Classify one SOW with the Excel-trained model")
    sow.add_argument("--category", required=True, choices=["Tower", "Ground"])
    sow.add_argument("--text", default="")
    sow.add_argument("--text-file", type=Path)

    sub.add_parser("run-samples", help="Build the Kirke workbook from the sample files")
    sub.add_parser("inspect", help="Show how the spreadsheets and MOPs in this folder pair")

    args = parser.parse_args(argv)
    settings = load_settings(args.root, args.config)
    try:
        if args.command == "classify-email":
            payload = _classify(settings, args)
        elif args.command == "classify-sow":
            payload = _classify_sow(settings, args)
        elif args.command == "process":
            payload = process_package(settings, _package(args))
        elif args.command == "run-samples":
            payload = run_samples(settings)
        else:
            payload = inspect_workspace(settings)
    except Exception as exc:
        print(json.dumps({"processingStatus": "failed", "emailStatus": "failed", "reason": str(exc)}))
        return 1
    print(json.dumps(payload, indent=2))
    return 0


def _classify(settings, args) -> dict:
    body = _body(args)
    tower = _split(args.tower) or settings.tower_recipients
    ground = _split(args.ground) or settings.ground_recipients
    return classify_email(
        {"emailId": args.email_id, "body": body, "from": args.sender},
        tower,
        ground,
        tower_senders=settings.tower_senders,
        ground_senders=settings.ground_senders,
    )


def _classify_sow(settings, args) -> dict:
    text = args.text
    if args.text_file:
        text = args.text_file.read_text(encoding="utf-8")
    model = SowClassifier.load(settings.path(settings.catalog_path))
    allowed = settings.request_types.get(args.category) or []
    prediction = model.predict(args.category, text, allowed) or {
        "requestType": "",
        "confidence": 0,
        "reason": "No similar SOW was found on that sheet",
    }
    return {
        "category": args.category,
        "requestType": prediction.get("requestType") or "",
        "confidence": prediction.get("confidence", 0),
        "reason": prediction.get("reason", ""),
        "sheetRanking": model.rank_sheets(text),
    }


def _package(args) -> dict:
    return {
        "emailId": args.email_id,
        "category": args.category,
        "originalTo": args.original_to,
        "originalFrom": args.original_from,
        "originalDate": args.original_date,
        "reason": args.reason,
        "body": _body(args),
        "spreadsheet": args.spreadsheet,
        "mops": args.mop,
    }


def _body(args) -> str:
    if getattr(args, "body_b64", ""):
        return base64.b64decode(args.body_b64).decode("utf-8")
    if getattr(args, "body_file", None):
        return args.body_file.read_text(encoding="utf-8")
    return ""


def _split(value: str) -> list[str]:
    return [part.strip() for part in value.replace(";", ",").split(",") if part.strip()]


if __name__ == "__main__":
    sys.exit(main())
