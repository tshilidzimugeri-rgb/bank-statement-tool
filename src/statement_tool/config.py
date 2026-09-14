"""Loading of the small YAML config files and .env settings."""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

import yaml
from dotenv import load_dotenv

PROJECT_ROOT = Path(__file__).resolve().parents[2]


@dataclass
class ClientRule:
    name: str
    match: list[str] = field(default_factory=list)


@dataclass
class BankLayout:
    key: str
    display_name: str
    detect: list[str] = field(default_factory=list)
    period_patterns: list[str] = field(default_factory=list)
    amount_style: str = "split"
    columns: dict[str, list[str]] = field(default_factory=dict)


@dataclass
class Settings:
    clients_config: Path
    banks_config: Path
    processed_db: Path
    output_workbook: Path
    incoming_pdfs_dir: Path
    gmail_client_secret_file: Path
    gmail_token_file: Path
    tesseract_cmd: str | None
    poppler_path: str | None


def load_settings(dotenv_path: Path | None = None) -> Settings:
    load_dotenv(dotenv_path or (PROJECT_ROOT / ".env"))

    def _path(env_name: str, default: str) -> Path:
        value = os.environ.get(env_name, default)
        p = Path(value)
        return p if p.is_absolute() else PROJECT_ROOT / p

    return Settings(
        clients_config=_path("CLIENTS_CONFIG", "config/clients.yaml"),
        banks_config=_path("BANKS_CONFIG", "config/banks.yaml"),
        processed_db=_path("PROCESSED_DB", "data/processed/processed.db"),
        output_workbook=_path("OUTPUT_WORKBOOK", "output/combined_statements.xlsx"),
        incoming_pdfs_dir=_path("INCOMING_PDFS_DIR", "data/incoming_pdfs"),
        gmail_client_secret_file=_path("GMAIL_CLIENT_SECRET_FILE", "credentials.json"),
        gmail_token_file=_path("GMAIL_TOKEN_FILE", "token.json"),
        tesseract_cmd=os.environ.get("TESSERACT_CMD"),
        poppler_path=os.environ.get("POPPLER_PATH"),
    )


def load_client_rules(path: Path) -> list[ClientRule]:
    if not path.exists():
        return []
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    rules = []
    for entry in data.get("clients", []) or []:
        rules.append(ClientRule(name=entry["name"], match=[m.lower() for m in entry.get("match", [])]))
    return rules


def load_bank_layouts(path: Path) -> tuple[dict[str, BankLayout], BankLayout]:
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    layouts: dict[str, BankLayout] = {}
    for key, entry in (data.get("banks") or {}).items():
        layouts[key] = BankLayout(
            key=key,
            display_name=entry.get("display_name", key),
            detect=[d.lower() for d in entry.get("detect", [])],
            period_patterns=entry.get("period_patterns", []),
            amount_style=entry.get("amount_style", "split"),
            columns=entry.get("columns", {}),
        )
    generic_entry = data.get("generic") or {}
    generic = BankLayout(
        key="generic",
        display_name=generic_entry.get("display_name", "Unknown/Generic"),
        detect=[],
        period_patterns=generic_entry.get("period_patterns", []),
        amount_style=generic_entry.get("amount_style", "split"),
        columns=generic_entry.get("columns", {}),
    )
    return layouts, generic
