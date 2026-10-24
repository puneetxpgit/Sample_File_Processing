"""Rebuild workflow/outlook-mop-automation.json from the classifier script."""

from __future__ import annotations

import json
from pathlib import Path
from uuid import uuid4

ROOT = Path(__file__).resolve().parents[1]


def main() -> None:
    js = (ROOT / "processor" / "n8n_classify.js").read_text(encoding="utf-8")
    sample_body = """Please process.

-----Original Message-----
From: Jane GC <jane.gc@vendor.com>
Sent: Tuesday, October 20, 2026 8:14 AM
To: Tower MOP <tower-mop@example.com>
Subject: GILLETT

Schedule attached.
"""
    sheet = str(ROOT / "GILLETT - Kirke Schedule Request 10.22.26.xlsx")
    mop = str(ROOT / "GILLETT - MOP DAY 2.docx")
    command = (
        f'python -m processor process --root "{ROOT}" '
        '--email-id "{{ $json.emailId }}" '
        '--category "{{ $json.category }}" '
        '--original-to "{{ $json.originalTo }}" '
        '--original-from "{{ $json.originalFrom }}" '
        '--original-date "{{ $json.originalDate }}" '
        '--reason "{{ $json.reason }}" '
        '--body-b64 "{{ $json.bodyBase64 }}" '
        '--spreadsheet "{{ $json.spreadsheetPath }}" '
        '--mop "{{ $json.mopPaths }}"'
    )

    def nid() -> str:
        return str(uuid4())

    ids = {name: nid() for name in ("manual", "config", "classify", "switch", "process", "note")}
    workflow = {
        "name": "Outlook MOP Automation",
        "nodes": [
            {
                "parameters": {},
                "id": ids["manual"],
                "name": "Manual Trigger",
                "type": "n8n-nodes-base.manualTrigger",
                "typeVersion": 1,
                "position": [260, 300],
            },
            {
                "parameters": {
                    "assignments": {
                        "assignments": [
                            {"id": nid(), "name": "towerRecipients", "value": "tower-mop@example.com", "type": "string"},
                            {"id": nid(), "name": "groundRecipients", "value": "ground-mop@example.com", "type": "string"},
                            {"id": nid(), "name": "towerSenders", "value": "ctx_tower_kirke@yahoo.com", "type": "string"},
                            {"id": nid(), "name": "groundSenders", "value": "ctx_ground_kirke@yahoo.com", "type": "string"},
                            {"id": nid(), "name": "emailId", "value": "sample-gillett", "type": "string"},
                            {"id": nid(), "name": "sampleBody", "value": sample_body, "type": "string"},
                            {"id": nid(), "name": "spreadsheetPath", "value": sheet, "type": "string"},
                            {"id": nid(), "name": "mopPaths", "value": mop, "type": "string"},
                        ]
                    },
                    "options": {},
                },
                "id": ids["config"],
                "name": "Classification Config",
                "type": "n8n-nodes-base.set",
                "typeVersion": 3.4,
                "position": [520, 300],
                "notes": "Replace example.com with the real Tower and Ground lists. Replace the sample body and paths with the Outlook message and saved attachments.",
            },
            {
                "parameters": {"mode": "runOnceForEachItem", "jsCode": js},
                "id": ids["classify"],
                "name": "Classify Forwarded Email",
                "type": "n8n-nodes-base.code",
                "typeVersion": 2,
                "position": [780, 300],
            },
            {
                "parameters": {
                    "rules": {
                        "values": [
                            _rule(nid(), "Tower"),
                            _rule(nid(), "Ground"),
                        ]
                    },
                    "options": {"fallbackOutput": "extra", "renameFallbackOutput": "Review"},
                },
                "id": ids["switch"],
                "name": "Route Category",
                "type": "n8n-nodes-base.switch",
                "typeVersion": 3.2,
                "position": [1040, 300],
            },
            {
                "parameters": {"command": command},
                "id": ids["process"],
                "name": "Process Excel and MOP",
                "type": "n8n-nodes-base.executeCommand",
                "typeVersion": 1,
                "position": [1320, 300],
                "notes": "Tower and Ground update their workbooks. Review writes output/Review and leaves the email uncompleted.",
            },
            {
                "parameters": {
                    "content": "Step 4 is Classify Forwarded Email. Step 5 is Process Excel and MOP. All three switch outputs run that step. The category decides which folder is written.",
                    "height": 160,
                    "width": 280,
                },
                "id": ids["note"],
                "name": "Sticky Note",
                "type": "n8n-nodes-base.stickyNote",
                "typeVersion": 1,
                "position": [240, 80],
            },
        ],
        "connections": {
            "Manual Trigger": {"main": [[{"node": "Classification Config", "type": "main", "index": 0}]]},
            "Classification Config": {"main": [[{"node": "Classify Forwarded Email", "type": "main", "index": 0}]]},
            "Classify Forwarded Email": {"main": [[{"node": "Route Category", "type": "main", "index": 0}]]},
            "Route Category": {
                "main": [
                    [{"node": "Process Excel and MOP", "type": "main", "index": 0}],
                    [{"node": "Process Excel and MOP", "type": "main", "index": 0}],
                    [{"node": "Process Excel and MOP", "type": "main", "index": 0}],
                ]
            },
        },
        "active": False,
        "settings": {"executionOrder": "v1"},
        "pinData": {},
        "meta": {"templateCredsSetupCompleted": True},
        "tags": [],
    }
    destination = ROOT / "workflow" / "outlook-mop-automation.json"
    destination.write_text(json.dumps(workflow, indent=2), encoding="utf-8")
    print(destination)


def _rule(rule_id: str, category: str) -> dict:
    return {
        "conditions": {
            "options": {"caseSensitive": True, "leftValue": "", "typeValidation": "strict", "version": 2},
            "conditions": [
                {
                    "id": rule_id,
                    "leftValue": "={{ $json.category }}",
                    "rightValue": category,
                    "operator": {"type": "string", "operation": "equals"},
                }
            ],
            "combinator": "and",
        },
        "renameOutput": True,
        "outputKey": category,
    }


if __name__ == "__main__":
    main()
