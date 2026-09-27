"""Made-up statements in layouts no bank config describes (see
fixtures/synthetic.py): each must be read exactly, and nothing wrong may ever
be APPROVED."""
import random
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent / "fixtures"))
from synthetic import generate, render_pdf  # noqa: E402

from statement_tool import config as config_mod  # noqa: E402
from statement_tool.extract.document import parse_document  # noqa: E402

PROJECT_ROOT = Path(__file__).parent.parent
SETTINGS = config_mod.load_settings(dotenv_path=PROJECT_ROOT / ".env.example")
LAYOUTS, GENERIC = config_mod.load_bank_layouts(PROJECT_ROOT / "config" / "banks.yaml")


def _read(path):
    results = parse_document(path, layouts=LAYOUTS, generic=GENERIC, client_rules=[], settings=SETTINGS,
                             interactive=False)
    return results, [t for r in results for t in r.transactions]


@pytest.mark.parametrize("seed", range(12))
def test_unseen_layout_read_exactly_and_approved(seed, tmp_path):
    statement = generate(seed)
    results, rows = _read(render_pdf(statement, tmp_path / "s.pdf"))
    assert [(t.date, t.debit, t.credit, t.balance) for t in rows] == [
        (w.date, w.debit, w.credit, w.balance) for w in statement.truth]
    assert results[0].account_number == statement.account
    if statement.style.closing_label:
        assert [r.status for r in results] == ["APPROVED"]
    else:
        # Nothing printed proves no rows are missing at the end.
        assert [r.status for r in results] == ["REVIEW_REQUIRED"]
        assert any("no closing balance or statement totals" in p for p in results[0].problems)
        assert {t.status for t in rows} == {"APPROVED"}  # every row read is itself confirmed


@pytest.mark.parametrize("seed", range(6))
def test_a_misprinted_amount_is_never_approved(seed, tmp_path):
    statement = generate(seed)
    rng = random.Random(seed)
    # Change one printed amount (not its balance), as a misprint or misread would.
    # Only transaction rows (date, description, amount, ... balance), and only
    # an amount that has a 1 or 2 to change - not a list line or a "0.00".
    candidates = [(pg, i) for pg, page in enumerate(statement.lines_by_page) for i, line in enumerate(page)
                  if len(line.split("   ")) >= 4 and any(d in line.split("   ")[2] for d in "12")
                  and line.split("   ")[2].strip("R()-CrD").replace(",", "").replace(" ", "").replace(".", "")
                  .isdigit() and float("0" + "".join(ch for ch in line.split("   ")[2] if ch.isdigit())) > 0]
    page, i = rng.choice(candidates)
    parts = statement.lines_by_page[page][i].split("   ")
    parts[2] = parts[2].replace("1", "7", 1) if "1" in parts[2] else parts[2].replace("2", "3", 1)
    statement.lines_by_page[page][i] = "   ".join(parts)
    results, rows = _read(render_pdf(statement, tmp_path / "s.pdf"))
    truth = {(w.date, w.debit, w.credit, w.balance) for w in statement.truth}
    wrong_but_approved = [t for t in rows if t.status == "APPROVED"
                          and (t.date, t.debit, t.credit, t.balance) not in truth]
    assert wrong_but_approved == []
    assert any(r.status == "REVIEW_REQUIRED" for r in results)
