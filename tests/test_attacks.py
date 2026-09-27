"""Adversarial tests: corrupt made-up statements (fixtures/attacks.py) and
require that none is APPROVED. A statement that says something false about
itself - a changed amount, balance or date, a missing, duplicated or
reordered row, a missing page, an OCR-style character - must come out
REVIEW_REQUIRED."""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent / "fixtures"))
from attacks import CAUGHT_WITH_TOTALS, MUST_BE_CAUGHT, mutations  # noqa: E402
from synthetic import generate, render_pdf  # noqa: E402

from statement_tool import config as config_mod  # noqa: E402
from statement_tool.extract.document import parse_document  # noqa: E402

PROJECT_ROOT = Path(__file__).parent.parent
SETTINGS = config_mod.load_settings(dotenv_path=PROJECT_ROOT / ".env.example")
LAYOUTS, GENERIC = config_mod.load_bank_layouts(PROJECT_ROOT / "config" / "banks.yaml")
SEEDS = range(4)


def _statuses(statement, path):
    results = parse_document(render_pdf(statement, path), layouts=LAYOUTS, generic=GENERIC, client_rules=[],
                             settings=SETTINGS, interactive=False)
    return [r.status if r.ok else "UNREADABLE" for r in results]


@pytest.mark.parametrize("seed", SEEDS)
def test_no_corruption_is_approved(seed, tmp_path):
    original = generate(seed)
    approved = []
    for name, corrupted in mutations(original):
        if corrupted is None:
            continue
        must_catch = name in MUST_BE_CAUGHT or (name in CAUGHT_WITH_TOTALS and original.recipe.totals)
        if not must_catch:
            continue
        if _statuses(corrupted, tmp_path / f"{name}.pdf") == ["APPROVED"]:
            approved.append(name)
    assert approved == [], f"seed {seed}: corrupted statements APPROVED: {approved}"


@pytest.mark.parametrize("seed", SEEDS)
def test_the_uncorrupted_statement_is_not_rejected(seed, tmp_path):
    """The other half: the same checks mustn't flag an honest statement that
    prints a closing balance or totals."""
    original = generate(seed)
    statuses = _statuses(original, tmp_path / "original.pdf")
    if original.style.closing_label or original.recipe.totals:
        assert statuses == ["APPROVED"]
    else:
        assert statuses == ["REVIEW_REQUIRED"]  # nothing printed proves the end is complete


def test_one_digit_change_is_caught_and_kept_as_printed(tmp_path):
    """Seed 2021 prints an "Upcoming debit orders" list above the opening
    balance "1 321.75", and a forged first amount (12 773.21 -> 13 773.21)
    also fits the balances if the opening is read as 321.75 with the "1"
    dropped. That reading must not win, and the flagged row must keep the
    whole figure as printed."""
    original = generate(2021)
    forged = dict(mutations(original))["one_digit_large_amount"]
    result = parse_document(render_pdf(forged, tmp_path / "forged.pdf"), layouts=LAYOUTS, generic=GENERIC,
                            client_rules=[], settings=SETTINGS, interactive=False)[0]
    assert result.status == "REVIEW_REQUIRED"
    assert result.report["opening_balance"] == 1321.75
    flagged = [t for t in result.transactions if t.check]
    assert len(flagged) == 1
    assert 13773.21 in (flagged[0].debit, flagged[0].credit, flagged[0].unassigned)
