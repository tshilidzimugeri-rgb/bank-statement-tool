"""Assigns each transaction a reporting category from config/categories.yaml."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import yaml

CATEGORY_TYPES = ("income", "expense", "transfer", "drawings")

UNCATEGORISED_INCOME = "Other income (uncategorised)"
UNCATEGORISED_EXPENSE = "Other expenses (uncategorised)"


@dataclass
class Category:
    name: str
    type: str
    match: list[str] = field(default_factory=list)
    # Whether amounts in this category include 15% VAT, by default. Can be
    # overridden per transaction in the workbook's VAT column.
    vat: bool = False


def load_categories(path: Path) -> list[Category]:
    if not path.exists():
        return []
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    categories = []
    for entry in data.get("categories", []) or []:
        ctype = str(entry.get("type", "")).strip().lower()
        if ctype not in CATEGORY_TYPES:
            raise ValueError(
                f"{path}: category {entry.get('name')!r} has type {ctype!r}; "
                f"expected one of {', '.join(CATEGORY_TYPES)}"
            )
        categories.append(
            Category(
                name=str(entry["name"]),
                type=ctype,
                match=[str(m).lower() for m in entry.get("match", []) or []],
                # Income and expenses are VAT-inclusive unless marked
                # otherwise; transfers and drawings never carry VAT.
                vat=bool(entry.get("vat", ctype in ("income", "expense"))),
            )
        )
    return categories


def with_fallbacks(categories: list[Category]) -> list[Category]:
    """The configured categories plus the two catch-alls, which the reports
    always list so uncategorised money is visible rather than dropped.
    """
    names = {c.name for c in categories}
    # VAT defaults to No until someone says what these are.
    extra = [
        Category(UNCATEGORISED_INCOME, "income", vat=False),
        Category(UNCATEGORISED_EXPENSE, "expense", vat=False),
    ]
    return categories + [c for c in extra if c.name not in names]


def categorize(description: str, is_credit: bool, categories: list[Category]) -> str:
    text = (description or "").lower()
    for category in categories:
        if any(needle in text for needle in category.match):
            return category.name
    return UNCATEGORISED_INCOME if is_credit else UNCATEGORISED_EXPENSE
