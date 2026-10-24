from pathlib import Path

import base64

import pytest
from fastapi.testclient import TestClient

from processor.api import create_app
from processor.settings import load_settings

ROOT = Path(__file__).resolve().parents[1]
TOKEN = "test-token"
BODY = """-----Original Message-----
From: Jane GC <jane.gc@vendor.com>
Sent: Tuesday, October 20, 2026 8:14 AM
To: Tower MOP <tower-mop@example.com>
Subject: GILLETT

Schedule attached.
"""


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setenv("MOP_API_TOKEN", TOKEN)
    settings = load_settings(ROOT)
    settings.output_dir = tmp_path / "out"
    return TestClient(create_app(settings))


def test_health_does_not_need_a_token(client):
    response = client.get("/health")
    assert response.status_code == 200
    assert response.json() == {"status": "ok"}


def test_classify_rejects_a_missing_token(client, monkeypatch):
    monkeypatch.delenv("MOP_API_TOKEN", raising=False)
    response = client.post("/classify", json={"body": BODY})
    assert response.status_code == 503

    monkeypatch.setenv("MOP_API_TOKEN", TOKEN)
    response = client.post("/classify", json={"body": BODY})
    assert response.status_code == 401


def test_classify_uses_the_inner_to_line(client):
    response = client.post(
        "/classify",
        headers={"Authorization": f"Bearer {TOKEN}"},
        json={
            "emailId": "gillett-test",
            "body": BODY,
            "towerRecipients": ["tower-mop@example.com"],
            "groundRecipients": ["ground-mop@example.com"],
        },
    )
    assert response.status_code == 200
    payload = response.json()
    assert payload["category"] == "Tower"
    assert payload["toSource"] == "innermost"
    assert payload["originalFrom"] == "jane.gc@vendor.com"
    assert payload["originalTo"] == "tower-mop@example.com"


def test_classify_gmail_uses_the_sender(client):
    response = client.post(
        "/classify",
        headers={"Authorization": f"Bearer {TOKEN}"},
        json={
            "emailId": "gmail-tower",
            "fromAddress": "ctx_tower_kirke@yahoo.com",
            "body": "Schedule attached.",
        },
    )
    assert response.status_code == 200
    payload = response.json()
    assert payload["category"] == "Tower"
    assert payload["toSource"] == "from"
    assert payload["originalFrom"] == "ctx_tower_kirke@yahoo.com"


def test_process_accepts_the_sheet_and_word_file(client):
    sheet = ROOT / "GILLETT - Kirke Schedule Request 10.22.26.xlsx"
    mop = ROOT / "GILLETT - MOP DAY 2.docx"
    with sheet.open("rb") as sheet_handle, mop.open("rb") as mop_handle:
        response = client.post(
            "/process",
            headers={"Authorization": f"Bearer {TOKEN}"},
            data={
                "emailId": "gillett-test",
                "category": "Tower",
                "originalTo": "tower-mop@example.com",
                "originalFrom": "jane.gc@vendor.com",
                "originalDate": "Tuesday, October 20, 2026 8:14 AM",
            },
            files={
                "spreadsheet": (sheet.name, sheet_handle, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                "mop": (mop.name, mop_handle, "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
            },
        )
    assert response.status_code == 200
    payload = response.json()
    assert payload["processingStatus"] == "completed"
    assert payload["invalidRows"] == 0
    assert payload["category"] == "Tower"
    assert sorted(row["mopFileName"] for row in payload["rows"]) == [
        "17548835_2026-10-21.docx",
        "17548835_2026-10-22.docx",
        "17548835_2026-10-22.docx",
    ]
    assert Path(payload["outputs"]["workbook"]).exists()
    workbook = payload["files"]["workbook"]
    assert workbook["fileName"] == "Tower_17548835_2026-10-20.xlsx"
    assert base64.b64decode(workbook["contentBase64"])[:2] == b"PK"
    assert [item["fileName"] for item in payload["files"]["mops"]] == [
        "17548835_2026-10-21.docx",
        "17548835_2026-10-22.docx",
    ]


def test_process_rejects_the_wrong_file_type(client, tmp_path):
    wrong = tmp_path / "notes.txt"
    wrong.write_text("nope", encoding="utf-8")
    with wrong.open("rb") as handle:
        response = client.post(
            "/process",
            headers={"Authorization": f"Bearer {TOKEN}"},
            data={"emailId": "bad", "category": "Tower"},
            files={"spreadsheet": ("notes.txt", handle, "text/plain")},
        )
    assert response.status_code == 400
