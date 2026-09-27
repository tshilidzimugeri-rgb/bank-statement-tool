"""3-, 6- and 12-month statements. Often these are monthly statements put
together, each month with its own opening and closing balance and totals;
or one long list whose dates ("11 Dec", "05 Jan") carry no year across New
Year. Read exactly, and still caught when something in them is wrong."""
import random
import sys
from datetime import date, timedelta
from pathlib import Path
from types import SimpleNamespace

import pytest

sys.path.insert(0, str(Path(__file__).parent / "fixtures"))
from synthetic import render_pdf  # noqa: E402

from statement_tool import config as config_mod  # noqa: E402
from statement_tool.extract.document import parse_document  # noqa: E402

PROJECT_ROOT = Path(__file__).parent.parent
SETTINGS = config_mod.load_settings(dotenv_path=PROJECT_ROOT / ".env.example")
LAYOUTS, GENERIC = config_mod.load_bank_layouts(PROJECT_ROOT / "config" / "banks.yaml")
DESCRIPTIONS = ["POS PURCHASE SHOPRITE", "DEBIT ORDER INSURANCE", "PAYMENT FROM CLIENT", "ATM WITHDRAWAL",
                "TRANSFER TO SAVINGS", "SALARY"]


def _money(v, cr):
    if cr:  # credits and positive balances marked "Cr", the rest unmarked
        return f"{abs(v):,.2f}" + ("Cr" if v >= 0 else "")
    return ("-" if v < 0 else "") + f"{abs(v):,.2f}"


def _months(start, end):
    m = date(start.year, start.month, 1)
    while m <= end:
        nxt = (m + timedelta(days=32)).replace(day=1)
        yield max(m, start), min(nxt - timedelta(days=1), end)
        m = nxt


def build(start, end, *, sections, date_fmt, cr=False, grand_totals=False, seed=0):
    """(printed lines, true transactions as (iso date, signed amount))."""
    rng = random.Random(seed)
    opening = 12345.67
    rows, d, bal = [], start - timedelta(days=1), opening
    while True:
        d += timedelta(days=rng.randint(1, 4))
        if d > end:
            break
        amount = round(rng.choice([-1, -1, 1]) * rng.uniform(20, 9000), 2)
        bal = round(bal + amount, 2)
        rows.append((d, rng.choice(DESCRIPTIONS), amount, bal))

    def row_line(r):
        return f"{r[0].strftime(date_fmt)}   {r[1]}   {_money(r[2], cr)}   {_money(r[3], cr)}"

    def totals(part):
        return [f"Total debits {-sum(a for _, _, a, _ in part if a < 0):,.2f}",
                f"Total credits {sum(a for _, _, a, _ in part if a > 0):,.2f}"]

    lines = ["Test Bank of Africa", "Account Number: 9900112233",
             f"Statement Period: {start:%d %b %Y} to {end:%d %b %Y}"]
    if grand_totals:
        lines += totals(rows)
    if sections:
        bal = opening
        for ms, me in _months(start, end):
            part = [r for r in rows if ms <= r[0] <= me]
            lines += [f"{ms:%B %Y}", f"Opening Balance {_money(bal, cr)}", "Date   Description   Amount   Balance"]
            lines += [row_line(r) for r in part]
            bal = part[-1][3] if part else bal
            lines += [f"Closing Balance {_money(bal, cr)}", *totals(part)]
    else:
        lines += [f"Opening Balance {_money(opening, cr)}", "Date   Description   Amount   Balance"]
        lines += [row_line(r) for r in rows]
        lines += [f"Closing Balance {_money(rows[-1][3], cr)}", *totals(rows)]
    return lines, [(r[0].isoformat(), r[2]) for r in rows], [row_line(r) for r in rows]


def read(lines, path):
    pages = [lines[i:i + 40] for i in range(0, len(lines), 40)]
    pages = [p + [f"Page {n + 1} of {len(pages)}"] for n, p in enumerate(pages)]
    return parse_document(render_pdf(SimpleNamespace(lines_by_page=pages), path), layouts=LAYOUTS, generic=GENERIC,
                          client_rules=[], settings=SETTINGS, interactive=False)


def _signed(t):
    return t.credit if t.credit else -t.debit if t.debit else None


SHAPES = {
    "3 months, monthly sections, no year on dates": dict(start=date(2025, 11, 1), end=date(2026, 1, 31),
                                                         sections=True, date_fmt="%d %b"),
    "3 months, one list over New Year, no year on dates": dict(start=date(2025, 11, 1), end=date(2026, 1, 31),
                                                               sections=False, date_fmt="%d %b"),
    "6 months, monthly sections": dict(start=date(2025, 9, 1), end=date(2026, 2, 28), sections=True,
                                       date_fmt="%d/%m/%Y"),
    "6 months, monthly sections and grand totals": dict(start=date(2025, 9, 1), end=date(2026, 2, 28),
                                                        sections=True, date_fmt="%d/%m/%Y", grand_totals=True),
    "6 months, monthly sections, Cr-marked": dict(start=date(2025, 9, 1), end=date(2026, 2, 28), sections=True,
                                                  date_fmt="%d %b", cr=True),
    "12 months, one list, no year on dates": dict(start=date(2025, 5, 1), end=date(2026, 4, 30), sections=False,
                                                  date_fmt="%d %b"),
}


@pytest.mark.parametrize("shape", SHAPES)
def test_multi_month_statement_read_exactly_and_approved(shape, tmp_path):
    lines, truth, _ = build(**SHAPES[shape], seed=7)
    results = read(lines, tmp_path / "s.pdf")
    assert [r.status for r in results] == ["APPROVED"], [r.problems for r in results]
    got = [(t.date, _signed(t)) for t in results[0].transactions]
    assert got == truth


def _corruptions(lines, row_lines):
    mid = lines.index(row_lines[len(row_lines) // 2])
    last = lines.index(row_lines[-1])
    first_total = next(i for i, ln in enumerate(lines) if ln.startswith("Total debits"))

    def replaced(i, text):
        return lines[:i] + [text] + lines[i + 1:]

    parts = lines[mid].split("   ")
    changed_amount = "   ".join(parts[:2] + [parts[2].replace(parts[2][-4], str((int(parts[2][-4]) + 1) % 10), 1)]
                                + parts[3:])
    yield "a row deleted", lines[:mid] + lines[mid + 1:]
    yield "a row duplicated", lines[:mid + 1] + [lines[mid]] + lines[mid + 1:]
    yield "an amount changed", replaced(mid, changed_amount)
    yield "the last row deleted", lines[:last] + lines[last + 1:]
    yield "one month's total changed", replaced(first_total, "Total debits 1.00")
    # A whole month's rows missing, its opening, closing and totals left in.
    month2 = [i for i, ln in enumerate(lines) if ln in row_lines][len(row_lines) // 3:2 * len(row_lines) // 3]
    yield "a block of rows missing", [ln for i, ln in enumerate(lines) if i not in set(month2)]


@pytest.mark.parametrize("shape", ["3 months, monthly sections, no year on dates", "6 months, monthly sections"])
def test_corrupted_multi_month_statement_is_not_approved(shape, tmp_path):
    lines, _, row_lines = build(**SHAPES[shape], seed=7)
    approved = [name for name, corrupted in _corruptions(lines, row_lines)
                if [r.status for r in read(corrupted, tmp_path / "c.pdf")] == ["APPROVED"]]
    assert approved == []
