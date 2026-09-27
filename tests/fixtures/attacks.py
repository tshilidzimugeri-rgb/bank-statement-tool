"""Ways of corrupting a made-up statement (see synthetic.py), for testing that
no corruption is ever APPROVED. Each yields (attack name, corrupted Statement).

Corruptions of what a statement says (amounts, balances, dates, missing,
duplicated or reordered rows, missing pages, OCR-style characters) must be
caught. "forge_*_full" and "meaning_change_*" rewrite the statement so all
its own arithmetic still holds - no check inside one statement can see
those (see README "What the checks cannot catch").
"""
import copy
import dataclasses
import re
from datetime import timedelta

from synthetic import _fmt_number, build

# Corruptions a statement's own figures must reveal.
MUST_BE_CAUGHT = {
    "delete_middle", "delete_first", "delete_last", "duplicate_row", "amount_x10", "amount_div10",
    "amount_one_decimal", "amount_plus_8c", "amount_first_row_x10", "amount_last_row_x10", "balance_plus_1",
    "opening_plus_10", "closing_plus_10", "swap_rows", "debit_credit_flip", "date_other_month",
    "date_other_year", "remove_middle_page", "remove_last_page", "truncate_end", "forge_duplicate_partial",
    "shift_into_next_month", "duplicate_altered_description", "one_digit_large_amount",
    "ocr_O_for_0_in_amount", "ocr_O_for_0_in_balance", "ocr_O_for_0_in_date", "ocr_I_for_1_in_amount",
    "ocr_I_for_1_in_balance", "ocr_I_for_1_in_date", "ocr_S_for_5_in_amount", "ocr_S_for_5_in_balance",
    "ocr_S_for_5_in_date",
}
# Caught only when the statement prints totals (the forger didn't update them).
CAUGHT_WITH_TOTALS = {"forge_offset_pair_partial"}


NUM = re.compile(r"\d{1,3}(?:[, ]\d{3})*\.\d{2}|\d+\.\d{2}")


def locate(st):
    """(page, index) of each table line."""
    where = []
    for text in st.table:
        for pg, page in enumerate(st.lines_by_page):
            if text in page:
                where.append((pg, page.index(text)))
                break
    return where


def change_first_number(text, style, fn, which=0):
    """Apply fn to the which-th money figure in a '   '-separated part."""
    matches = list(NUM.finditer(text))
    if len(matches) <= which:
        return None
    m = matches[which]
    value = float(m.group().replace(",", "").replace(" ", ""))
    new = fn(value)
    if isinstance(new, str):
        return text[:m.start()] + new + text[m.end():]
    return text[:m.start()] + _fmt_number(new, style.thousands) + text[m.end():]


def mutations(st, rng=None):
    where = locate(st)
    if len(where) < 4:
        return
    s = st.style
    mid = len(where) // 2

    def with_lines(fn):
        new = copy.deepcopy(st)
        fn(new.lines_by_page)
        return new

    def set_line(i, text):
        pg, idx = where[i]
        return lambda pages: pages[pg].__setitem__(idx, text)

    def money_part_index(line):
        return 2  # date, description, first money part

    def mutate_amount(i, fn):
        pg, idx = where[i]
        parts = st.lines_by_page[pg][idx].split("   ")
        k = money_part_index(parts)
        if s.amount_style == "split_zero":
            k = 2 if NUM.search(parts[2]) and float(NUM.search(parts[2]).group().replace(",", "").replace(" ", "")) else 3
        changed = change_first_number(parts[k], s, fn)
        if changed is None or changed == parts[k]:
            return None
        parts[k] = changed
        return with_lines(set_line(i, "   ".join(parts)))

    yield "delete_middle", with_lines(lambda p: p[where[mid][0]].pop(where[mid][1]))
    yield "delete_first", with_lines(lambda p: p[where[0][0]].pop(where[0][1]))
    yield "delete_last", with_lines(lambda p: p[where[-1][0]].pop(where[-1][1]))
    yield "duplicate_row", with_lines(lambda p: p[where[mid][0]].insert(where[mid][1] + 1, st.table[mid]))
    yield "amount_x10", mutate_amount(mid, lambda v: v * 10)
    yield "amount_div10", mutate_amount(mid, lambda v: round(v / 10, 2))
    yield "amount_one_decimal", mutate_amount(mid, lambda v: _fmt_number(v, s.thousands)[:-1])
    yield "amount_plus_8c", mutate_amount(mid, lambda v: v + 0.08)
    yield "amount_first_row_x10", mutate_amount(0, lambda v: v * 10)
    yield "amount_last_row_x10", mutate_amount(len(where) - 1, lambda v: v * 10)

    # balance of a row that shows one
    rows_with_balance = [i for i, t in enumerate(st.table) if len(t.split("   ")) >= 4]
    if rows_with_balance:
        i = rows_with_balance[len(rows_with_balance) // 2]
        pg, idx = where[i]
        parts = st.lines_by_page[pg][idx].split("   ")
        b = len(parts) - 1 - (1 if s.trailing_column and not re.search(r"Cr|-", parts[-1]) and len(parts) > 4 else 0)
        changed = change_first_number(parts[b], s, lambda v: v + 1.00)
        if changed:
            parts[b] = changed
            yield "balance_plus_1", with_lines(set_line(i, "   ".join(parts)))

    for label, key in (("opening_plus_10", s.opening_label), ("closing_plus_10", s.closing_label)):
        if not key:
            continue
        for pg, page in enumerate(st.lines_by_page):
            for idx, line in enumerate(page):
                if line.startswith(key):
                    changed = change_first_number(line, s, lambda v: v + 10)
                    yield label, with_lines(lambda p, pg=pg, idx=idx, changed=changed: p[pg].__setitem__(idx, changed))
                    break
            else:
                continue
            break

    # reorder two rows with different amounts on different days
    j = next((j for j in range(mid, len(where) - 1) if st.table[j].split("   ")[2] != st.table[j + 1].split("   ")[2]), None)
    if j is not None:
        a, b = st.table[j], st.table[j + 1]
        yield "swap_rows", with_lines(lambda p: (p[where[j][0]].__setitem__(where[j][1], b),
                                                 p[where[j + 1][0]].__setitem__(where[j + 1][1], a)))

    # debit <-> credit, where the text shows the direction
    pg, idx = where[mid]
    parts = st.lines_by_page[pg][idx].split("   ")
    flip = {"signed": lambda t: t[1:] if t.startswith("-") else "-" + t,
            "crdr": lambda t: t.replace("Dr", "Cr") if "Dr" in t else t.replace("Cr", "Dr"),
            "cr_only": lambda t: t[:-2] if t.endswith("Cr") else t + "Cr",
            "trailing_minus": lambda t: t[:-1] if t.endswith("-") else t + "-",
            "parens": lambda t: t[1:-1] if t.startswith("(") else f"({t})"}.get(s.amount_style)
    if flip:
        parts[2] = flip(parts[2])
        yield "debit_credit_flip", with_lines(set_line(mid, "   ".join(parts)))

    # dates: another month; another year
    date_text = st.table[mid].split("   ")[0]
    for label, fn in (("date_other_month", lambda d: re.sub(r"(?<=[/.-])(\d{2})(?=[/.-])", lambda m: f"{(int(m.group()) % 12) + 1:02d}", d, count=1)
                       if re.search(r"\d{2}[/.-]\d{2}[/.-]", d) else re.sub(r"Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec",
                       lambda m: {"Jan": "Jun", "Feb": "Jul", "Mar": "Aug", "Apr": "Sep", "May": "Oct", "Jun": "Nov", "Jul": "Dec",
                                  "Aug": "Jan", "Sep": "Feb", "Oct": "Mar", "Nov": "Apr", "Dec": "May"}[m.group()], d, count=1)),
                      ("date_other_year", lambda d: re.sub(r"20(2\d)", lambda m: f"20{int(m.group(1)) - 3}", d, count=1)
                       if re.search(r"20\d\d", d) else re.sub(r"(?<= )(\d\d)$", lambda m: f"{int(m.group()) - 3:02d}", d))):
        new_date = fn(date_text)
        if new_date != date_text:
            yield label, with_lines(set_line(mid, st.table[mid].replace(date_text, new_date, 1)))

    if len(st.lines_by_page) > 1:
        yield "remove_middle_page" if len(st.lines_by_page) > 2 else "remove_last_page", \
            with_lines(lambda p: p.pop(1 if len(p) > 2 else len(p) - 1))

    # truncate: the last third of the transactions, the closing line and the footer are gone
    cut = where[len(where) * 2 // 3]
    yield "truncate_end", with_lines(lambda p: (p.__delitem__(slice(cut[0] + 1, None)), p[cut[0]].__delitem__(slice(cut[1], None))))

    # --- reconciliation attacks ---------------------------------------------------
    rec = st.recipe
    rows = list(rec.day_rows)

    def forged(new_rows, keep_summary):
        new = build(dataclasses.replace(rec, day_rows=new_rows))
        if keep_summary:  # forger edited the transactions but not the summary lines
            summary = {l.split("   ")[0]: l for page in st.lines_by_page for l in page
                       if l.startswith((s.closing_label or "~~", "Total debits", "Total credits"))}
            for page in new.lines_by_page:
                for k, l in enumerate(page):
                    key = l.split("   ")[0]
                    if key in summary:
                        page[k] = summary[key]
        return new

    i = len(rows) // 2
    d, desc, amount, fee, accrued = rows[i]
    offset = rows[:i + 1] + [(d, "TRANSFER Loan REF A777", 5000.0, None, None),
                             (d, "TRANSFER Loan REF A778", -5000.0, None, None)] + rows[i + 1:]
    yield "forge_offset_pair_full", forged(offset, keep_summary=False)
    yield "forge_offset_pair_partial", forged(offset, keep_summary=True)

    debits = [k for k, r in enumerate(rows) if r[2] < -200]
    if len(debits) >= 2:
        a, b = debits[0], debits[-1]
        moved = list(rows)
        moved[a] = (rows[a][0], rows[a][1], round(rows[a][2] - 100, 2), rows[a][3], rows[a][4])
        moved[b] = (rows[b][0], rows[b][1], round(rows[b][2] + 100, 2), rows[b][3], rows[b][4])
        yield "forge_preserve_totals", forged(moved, keep_summary=False)

    dup = rows[:i + 1] + [rows[i]] + rows[i + 1:]
    yield "forge_duplicate_full", forged(dup, keep_summary=False)
    yield "forge_duplicate_partial", forged(dup, keep_summary=True)

    # shift the last transaction into the next month (order kept)
    last = rows[-1]
    shifted = rows[:-1] + [(rec.end + timedelta(days=1), last[1], last[2], last[3], last[4])]
    yield "shift_into_next_month", forged(shifted, keep_summary=False)

    # duplicate line with a slightly different description, balances untouched
    pg, idx = where[mid]
    parts = st.table[mid].split("   ")
    parts[1] = parts[1] + " X"
    yield "duplicate_altered_description", with_lines(lambda p: p[pg].insert(idx + 1, "   ".join(parts)))

    # one digit in a large amount
    big = next((k for k, t in enumerate(st.table) if (lambda m: m and float(m.group().replace(",", "").replace(" ", "")) >= 1000)(NUM.search(t.split("   ")[2]))), None)
    if big is not None:
        yield "one_digit_large_amount", mutate_amount(big, lambda v: v + 1000 if int(v / 1000) % 10 != 9 else v - 1000)

    # OCR-style character swaps
    for label, bad, good in (("ocr_O_for_0", "O", "0"), ("ocr_I_for_1", "I", "1"), ("ocr_S_for_5", "S", "5")):
        k = next((k for k, t in enumerate(st.table) if good in t.split("   ")[2]), None)
        if k is not None:
            parts = st.table[k].split("   ")
            parts[2] = parts[2].replace(good, bad, 1)
            yield label + "_in_amount", with_lines(set_line(k, "   ".join(parts)))
        k = next((k for k, t in enumerate(st.table) if len(t.split("   ")) >= 4 and good in t.split("   ")[-1]), None)
        if k is not None:
            parts = st.table[k].split("   ")
            parts[-1] = parts[-1].replace(good, bad, 1)
            yield label + "_in_balance", with_lines(set_line(k, "   ".join(parts)))
        k = next((k for k, t in enumerate(st.table) if good in t.split("   ")[0]), None)
        if k is not None:
            parts = st.table[k].split("   ")
            parts[0] = parts[0].replace(good, bad, 1)
            yield label + "_in_date", with_lines(set_line(k, "   ".join(parts)))

    # meaning changes that keep the arithmetic: a different description
    parts = st.table[mid].split("   ")
    parts[1] = "SALARY Loan Received"
    yield "meaning_change_description", with_lines(set_line(mid, "   ".join(parts)))

