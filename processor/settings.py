from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path


def project_root() -> Path:
    return Path(__file__).resolve().parents[1]


@dataclass
class Settings:
    root: Path
    tower_recipients: list[str] = field(default_factory=list)
    ground_recipients: list[str] = field(default_factory=list)
    tower_senders: list[str] = field(default_factory=list)
    ground_senders: list[str] = field(default_factory=list)
    template_path: Path = Path("Kirke Output format.xlsx")
    catalog_path: Path = Path("Kirke Output format (1).xlsx")
    output_dir: Path = Path("output")
    mop_date_format: str = "YYYY-MM-DD"
    brief_sow_word_limit: int = 500
    request_type_min_confidence: int = 90
    match_strategy: str = "both"
    populate_location_and_site_name: bool = True
    header_synonyms: dict[str, list[str]] = field(default_factory=dict)
    request_types: dict[str, list[str]] = field(default_factory=dict)
    explicit_phrases: dict[str, list] = field(default_factory=dict)
    sample_jobs: list[dict] = field(default_factory=list)
    ai_base_url: str = ""
    ai_api_key: str = ""
    ai_model: str = ""

    def path(self, value: str | Path) -> Path:
        path = Path(value)
        return path if path.is_absolute() else self.root / path


def load_settings(root: Path | None = None, config_path: Path | None = None) -> Settings:
    root = (root or project_root()).resolve()
    path = config_path or (root / "config" / "settings.json")
    data = json.loads(path.read_text(encoding="utf-8"))
    import os

    strategy = str(data.get("matchStrategy", "both"))
    if strategy not in {"both", "filename", "content"}:
        raise ValueError(f"matchStrategy must be both, filename, or content, not {strategy!r}")

    return Settings(
        root=root,
        tower_recipients=_list(data.get("towerRecipients")),
        ground_recipients=_list(data.get("groundRecipients")),
        tower_senders=_list(data.get("towerSenders")),
        ground_senders=_list(data.get("groundSenders")),
        template_path=Path(data.get("templatePath", "Kirke Output format.xlsx")),
        catalog_path=Path(data.get("catalogPath", "Kirke Output format (1).xlsx")),
        output_dir=Path(data.get("outputDir", "output")),
        mop_date_format=str(data.get("mopDateFormat", "YYYY-MM-DD")),
        brief_sow_word_limit=int(data.get("briefSowWordLimit", 500)),
        request_type_min_confidence=int(data.get("requestTypeMinConfidence", 90)),
        match_strategy=strategy,
        populate_location_and_site_name=bool(data.get("populateLocationAndSiteName", True)),
        header_synonyms={k: _list(v) for k, v in (data.get("headerSynonyms") or {}).items()},
        request_types={k: _list(v) for k, v in (data.get("requestTypes") or {}).items()},
        explicit_phrases=data.get("explicitPhrases") or {},
        sample_jobs=list(data.get("sampleJobs") or []),
        ai_base_url=os.environ.get("MOP_AI_BASE_URL", ""),
        ai_api_key=os.environ.get("MOP_AI_API_KEY", "") or os.environ.get("OPENAI_API_KEY", ""),
        ai_model=os.environ.get("MOP_AI_MODEL", "gpt-4o-mini"),
    )


def _list(value) -> list[str]:
    if value is None:
        return []
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item).strip()]
    return [part.strip() for part in str(value).replace(";", ",").split(",") if part.strip()]
