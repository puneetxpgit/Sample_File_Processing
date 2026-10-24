"""Local API for one forwarded email plus its spreadsheet and MOP files.

Run with MOP_API_TOKEN set:

    python -m processor.api

Classify and process both require ``Authorization: Bearer <token>``.
The server listens on 127.0.0.1:8000 unless MOP_API_HOST or MOP_API_PORT is set.
"""

from __future__ import annotations

import base64
import os
import re
from pathlib import Path

from fastapi import Depends, FastAPI, File, Form, Header, HTTPException, UploadFile
from pydantic import BaseModel, Field

from processor.classify_email import classify_email
from processor.pipeline import process_package
from processor.settings import Settings, load_settings

ALLOWED_SHEETS = {".xlsx", ".xlsm", ".csv", ".xls"}
ALLOWED_MOPS = {".docx"}


class ClassifyRequest(BaseModel):
    body: str = ""
    emailId: str = ""
    subject: str = ""
    fromAddress: str = ""
    towerRecipients: list[str] = Field(default_factory=list)
    groundRecipients: list[str] = Field(default_factory=list)
    towerSenders: list[str] = Field(default_factory=list)
    groundSenders: list[str] = Field(default_factory=list)


def create_app(settings: Settings | None = None) -> FastAPI:
    app = FastAPI(title="Outlook MOP automation")
    app.state.settings = settings or load_settings()

    @app.get("/health")
    def health() -> dict:
        return {"status": "ok"}

    @app.post("/classify")
    def classify(payload: ClassifyRequest, _: None = Depends(require_token)) -> dict:
        current = _settings(app)
        tower = payload.towerRecipients or current.tower_recipients
        ground = payload.groundRecipients or current.ground_recipients
        tower_senders = payload.towerSenders or current.tower_senders
        ground_senders = payload.groundSenders or current.ground_senders
        return classify_email(
            {
                "emailId": payload.emailId,
                "subject": payload.subject,
                "body": payload.body,
                "from": payload.fromAddress,
            },
            tower,
            ground,
            tower_senders=tower_senders,
            ground_senders=ground_senders,
        )

    @app.post("/process")
    def process(
        emailId: str = Form(...),
        category: str = Form(...),
        originalTo: str = Form(""),
        originalFrom: str = Form(""),
        originalDate: str = Form(""),
        reason: str = Form(""),
        body: str = Form(""),
        spreadsheet: UploadFile = File(...),
        mop: list[UploadFile] = File(default_factory=list),
        _: None = Depends(require_token),
    ) -> dict:
        if category not in {"Tower", "Ground", "Review"}:
            raise HTTPException(status_code=400, detail="category must be Tower, Ground, or Review")
        current = _settings(app)
        folder = _inbox(current, emailId)
        sheet_path = _store(folder, spreadsheet, ALLOWED_SHEETS, "spreadsheet")
        mop_paths = [_store(folder, item, ALLOWED_MOPS, "mop") for item in mop if item.filename]
        result = process_package(
            current,
            {
                "emailId": emailId,
                "category": category,
                "originalTo": originalTo,
                "originalFrom": originalFrom,
                "originalDate": originalDate,
                "reason": reason,
                "body": body,
                "spreadsheet": str(sheet_path),
                "mops": [str(path) for path in mop_paths],
            },
        )
        result["files"] = _response_files(result)
        return result

    return app


def require_token(authorization: str = Header(default="")) -> None:
    expected = os.environ.get("MOP_API_TOKEN", "").strip()
    if not expected:
        raise HTTPException(status_code=503, detail="MOP_API_TOKEN is not configured")
    if authorization != f"Bearer {expected}":
        raise HTTPException(status_code=401, detail="Unauthorized")


def _settings(app: FastAPI) -> Settings:
    return app.state.settings


def _inbox(settings: Settings, email_id: str) -> Path:
    folder = settings.path(settings.output_dir) / "inbox" / _safe_id(email_id)
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def _store(folder: Path, upload: UploadFile, allowed: set[str], label: str) -> Path:
    name = _safe_filename(upload.filename or "")
    suffix = Path(name).suffix.lower()
    if suffix not in allowed:
        raise HTTPException(status_code=400, detail=f"{label} must be one of: {', '.join(sorted(allowed))}")
    target = folder / name
    with target.open("wb") as handle:
        while True:
            chunk = upload.file.read(1024 * 1024)
            if not chunk:
                break
            handle.write(chunk)
    return target


def _safe_id(email_id: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", email_id).strip("._")
    return (cleaned or "email")[:80]


def _response_files(result: dict) -> dict:
    """Return this email's workbook and renamed MOPs. The category master stays on disk."""
    outputs = result.get("outputs") or {}
    names: list[str] = []
    seen: set[str] = set()
    for row in list(result.get("rows") or []) + list(result.get("reviewRows") or []):
        name = str(row.get("mopFileName") or "")
        if name and name not in seen:
            seen.add(name)
            names.append(name)
    saved = {
        Path(path).name: Path(path)
        for path in list(outputs.get("mops") or []) + list(outputs.get("reviewMops") or [])
    }
    workbook = outputs.get("replyWorkbook") or outputs.get("workbook")
    mops = []
    for name in names:
        if name not in saved:
            continue
        encoded = _encoded_file(saved[name])
        if encoded:
            mops.append(encoded)
    files = {
        "workbook": _encoded_file(Path(workbook)) if workbook else None,
        "mops": mops,
    }
    review = outputs.get("reviewReplyWorkbook") or ""
    if review:
        files["reviewWorkbook"] = _encoded_file(Path(review))
    return files


def _encoded_file(path: Path) -> dict | None:
    if not path.is_file():
        return None
    mime = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        if path.suffix.lower() == ".xlsx"
        else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    return {
        "fileName": path.name,
        "mimeType": mime,
        "contentBase64": base64.b64encode(path.read_bytes()).decode("ascii"),
    }


def _safe_filename(name: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._ -]+", "_", Path(name).name).strip(" .")
    return cleaned or "upload"


app = create_app()


def main() -> None:
    if not os.environ.get("MOP_API_TOKEN", "").strip():
        raise SystemExit("Set MOP_API_TOKEN before starting the API.")
    import uvicorn

    host = os.environ.get("MOP_API_HOST", "127.0.0.1")
    port = int(os.environ.get("MOP_API_PORT", "8000"))
    uvicorn.run(app, host=host, port=port)


if __name__ == "__main__":
    main()
