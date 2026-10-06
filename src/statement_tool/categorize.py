"""Assigns each transaction a reporting category from config/categories.yaml."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import yaml

CATEGORY_TYPES = ("income", "expense", "transfer", "drawings")

UNCATEGORISED_INCOME = "Other income (uncategorised)"
UNCATEGORISED_EXPENSE = "Other expenses (uncategorised)"


DIRECTIONS = ("in", "out", "any")


@dataclass
class Category:
    name: str
    type: str
    match: list[str] = field(default_factory=list)
    # Whether amounts in this category include 15% VAT, by default. Can be
    # overridden per transaction in the workbook's VAT column.
    vat: bool = False
    # Which money the rule applies to: "in", "out" or "any". Income is money
    # in and expenses money out by default, so a deposit whose description
    # mentions "municipal" isn't counted as a (negative) electricity expense.
    direction: str = "any"
    # The category's VAT follows the business's VAT registration (customer
    # payments include VAT only when the business is registered).
    vat_if_registered: bool = False


def default_direction(ctype: str) -> str:
    return {"income": "in", "expense": "out"}.get(ctype, "any")


def for_vat_registration(categories: list[Category], registered: bool) -> list[Category]:
    """The categories as they apply to a business that is - or isn't -
    VAT-registered: one that isn't charges and claims no VAT at all."""
    out = []
    for c in categories:
        vat = (c.vat or c.vat_if_registered) if registered else False
        out.append(Category(c.name, c.type, c.match, vat, c.direction, c.vat_if_registered))
    return out


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
        direction = str(entry.get("direction", default_direction(ctype))).strip().lower()
        if direction not in DIRECTIONS:
            raise ValueError(f"{path}: category {entry.get('name')!r} has direction {direction!r}; "
                             f"expected one of {', '.join(DIRECTIONS)}")
        vat_setting = entry.get("vat", ctype in ("income", "expense"))
        categories.append(
            Category(
                name=str(entry["name"]),
                type=ctype,
                match=[str(m).lower() for m in entry.get("match", []) or []],
                # Income and expenses are VAT-inclusive unless marked
                # otherwise; transfers and drawings never carry VAT.
                # "if registered": only when the business is VAT-registered.
                vat=vat_setting is True,
                direction=direction,
                vat_if_registered=str(vat_setting).strip().lower() == "if registered",
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
    way = "in" if is_credit else "out"
    for category in categories:
        if category.direction not in ("any", way):
            continue
        if any(needle in text for needle in category.match):
            return category.name
    return UNCATEGORISED_INCOME if is_credit else UNCATEGORISED_EXPENSE
