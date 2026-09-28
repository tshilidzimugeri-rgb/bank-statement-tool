"""Made-up bank statements in many layouts, with the right answer, to test
the reader on statement styles it has never seen.

Each statement varies: date format, how money in/out is shown (debit/credit
columns, minus signs, Cr/Dr, trailing minus, brackets), thousands separator,
fee and extra columns, balances on every row or only at day end, opening and
closing lines, page breaks with "brought forward" lines, and lists that
aren't transactions. Nothing here is a real bank's layout.

On request (generate's decimal and fee_line), figures are written with a
decimal comma ("1 234,56"), and each fee sits on its own line under its
entry, the date printed once - with fees totalled apart from the debits.
"""
from __future__ import annotations

import random
from dataclasses import dataclass, field
from datetime import date, timedelta
from pathlib import Path

DATE_FORMATS = ["dd/mm/yyyy", "dd-mm-yyyy", "dd.mm.yyyy", "yyyy-mm-dd", "dd Mon yyyy", "dd Mon yy", "dd Mon",
                "Mon dd, yyyy"]
AMOUNT_STYLES = ["split", "split_zero", "signed", "crdr", "cr_only", "trailing_minus", "parens"]
BANK_NAMES = ["Harbour Mutual Bank", "Karoo Savings Bank", "Protea Commercial", "Baobab Bank Ltd",
              "Summit Finance Bank", "Lowveld Trust Bank"]
WORDS = ["PAYMENT", "TRANSFER", "DEBIT ORDER", "POS PURCHASE", "SALARY", "DEPOSIT", "CASH WITHDRAWAL", "FEE",
         "INSURANCE", "ELECTRICITY", "WATER", "RENT", "GROCERIES", "FUEL", "AIRTIME", "INTEREST"]
REFS = ["REF A{n}", "CARD 4567*{n}", "INV {n}", "ACC {n}X", "{n}ZA", "TRN{n}"]
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


@dataclass
class Style:
    date_format: str
    amount_style: str
    thousands: str
    currency: bool
    opening_label: str | None
    closing_label: str | None
    fee_column: bool
    trailing_column: bool
    balance_every_row: bool
    summary_list: bool
    rows_per_page: int
    decimal: str = "."  # "," for "1 234,56"
    # Each fee on its own undated line under its entry, which then carries
    # the balance; an entry can be just a fee ("SMS NOTIFICATION").
    fee_line: bool = False


@dataclass
class TruthRow:
    date: str
    debit: float | None
    credit: float | None
    balance: float | None  # as printed; None where the layout prints none


@dataclass
class Statement:
    style: Style
    lines_by_page: list[list[str]]
    truth: list[TruthRow]
    account: str
    opening: float
    closing: float
    notes: list[str] = field(default_factory=list)
    table: list[str] = field(default_factory=list)  # the transaction lines, in order
    truth_rows_per_line: list[int] = field(default_factory=list)  # truth rows each table line produces


def random_style(rng: random.Random) -> Style:
    amount_style = rng.choice(AMOUNT_STYLES)
    return Style(
        date_format=rng.choice(DATE_FORMATS),
        amount_style=amount_style,
        thousands=rng.choice([",", ",", " ", ""]),
        currency=rng.random() < 0.2,
        opening_label=rng.choice(["Opening Balance", "Balance Brought Forward", "Previous Balance", "Opening Balance"]),
        closing_label=rng.choice(["Closing Balance", "Balance Carried Forward", None]),
        fee_column=rng.random() < 0.25,
        trailing_column=amount_style == "cr_only" and rng.random() < 0.5,
        balance_every_row=rng.random() < 0.85,
        summary_list=rng.random() < 0.3,
        rows_per_page=rng.choice([25, 40, 1000]),
    )


def _fmt_number(value: float, thousands: str, decimal: str = ".") -> str:
    whole, cents = f"{abs(value):.2f}".split(".")
    groups = []
    while len(whole) > 3:
        groups.insert(0, whole[-3:])
        whole = whole[:-3]
    groups.insert(0, whole)
    return f"{thousands.join(groups)}{decimal}{cents}"


def _fmt_date(d: date, fmt: str) -> str:
    mon = MONTHS[d.month - 1]
    return {
        "dd/mm/yyyy": f"{d.day:02d}/{d.month:02d}/{d.year}",
        "dd-mm-yyyy": f"{d.day:02d}-{d.month:02d}-{d.year}",
        "dd.mm.yyyy": f"{d.day:02d}.{d.month:02d}.{d.year}",
        "yyyy-mm-dd": d.isoformat(),
        "dd Mon yyyy": f"{d.day:02d} {mon} {d.year}",
        "dd Mon yy": f"{d.day:02d} {mon} {d.year % 100:02d}",
        "dd Mon": f"{d.day:02d} {mon}",
        "Mon dd, yyyy": f"{mon} {d.day:02d}, {d.year}",
    }[fmt]


def _money(value: float, style: Style, kind: str) -> str:
    """kind: 'debit', 'credit' (a movement) or 'balance'."""
    body = ("R" if style.currency else "") + _fmt_number(value, style.thousands, style.decimal)
    s = style.amount_style
    negative = value < 0 if kind == "balance" else kind == "debit"
    if s == "crdr":
        return body + ("Dr" if negative else "Cr")
    if s == "cr_only":
        return body if negative else body + "Cr"
    if s == "trailing_minus":
        return body + ("-" if negative else "")
    if s == "parens":
        return f"({body})" if negative else body
    if kind == "balance" or s == "signed":
        return ("-" if negative else "") + body
    return body  # split columns: the column shows the direction


@dataclass
class Recipe:
    """Everything a statement is made from - kept so a copy can be rebuilt
    with altered transactions and every balance recalculated (forgeries)."""
    style: Style
    start: date
    end: date
    account: str
    opening: float
    bank: str
    account_label: str
    summary_lines: list[str]
    totals: bool
    day_rows: list  # (date, description, signed amount, fee, accrued)


def generate(seed: int, decimal: str = ".", fee_line: bool = False) -> Statement:
    rng = random.Random(seed)
    style = random_style(rng)
    if decimal == ",":
        style.decimal = ","
        style.thousands = {",": "."}.get(style.thousands, style.thousands)
    if fee_line:
        style.fee_column = style.fee_line = style.balance_every_row = True
        if style.amount_style == "split_zero":
            style.amount_style = "split"
    start = date(rng.choice([2025, 2026]), rng.randint(1, 12), 1)
    end = (start + timedelta(days=32)).replace(day=1) - timedelta(days=1)
    account = str(rng.randint(10**9, 10**10 - 1))
    balance = round(rng.uniform(-2000, 30000), 2)
    if style.amount_style in ("split", "split_zero") and balance < 0:
        balance = abs(balance)

    day_rows = []
    day = start
    for _ in range(rng.randint(12, 55)):
        day = min(end, day + timedelta(days=rng.choice([0, 0, 1, 1, 2, 3])))
        if style.fee_line and rng.random() < 0.15:
            day_rows.append((day, "SMS NOTIFICATION", 0.0, 1.25, None))  # an entry that is only its fee
            continue
        credit = rng.random() < 0.3
        size = rng.choice([rng.uniform(1, 150), rng.uniform(100, 3000), rng.uniform(1000, 60000)])
        amount = round(size, 2) * (1 if credit else -1)
        desc = f"{rng.choice(WORDS)} {rng.choice(WORDS).title()} {rng.choice(REFS).format(n=rng.randint(100, 99999))}"
        fee = round(rng.choice([2.0, 5.5, 8.0, 12.5]), 2) if style.fee_column and rng.random() < 0.3 else None
        accrued = round(rng.choice([3.0, 8.0, 14.8]), 2) if style.trailing_column and rng.random() < 0.3 else None
        day_rows.append((day, desc, round(amount, 2), fee, accrued))

    bank = rng.choice(BANK_NAMES)
    account_label = rng.choice(["Account Number:", "Account No.", "Account number", "Account :"])
    summary = []
    if style.summary_list:
        summary = ["Upcoming debit orders"] + [
            f"{_fmt_date(start + timedelta(days=rng.randint(0, 27)), style.date_format)}   {rng.choice(WORDS)}   "
            f"{_money(round(rng.uniform(50, 900), 2), style, 'debit')}" for _ in range(3)]
    totals = rng.random() < 0.4
    return build(Recipe(style, start, end, account, balance, bank, account_label, summary, totals, day_rows))


def build(recipe: Recipe) -> Statement:
    style = recipe.style
    balance = recipe.opening
    day_rows = recipe.day_rows
    truth: list[TruthRow] = []
    table: list[str] = []
    rows_per_line: list[int] = []
    for i, (d, desc, amount, fee, accrued) in enumerate(day_rows):
        balance = round(balance + amount - (fee or 0), 2)
        last_of_day = i + 1 == len(day_rows) or day_rows[i + 1][0] != d
        show_balance = style.balance_every_row or last_of_day
        date_text = _fmt_date(d, style.date_format)
        if style.amount_style == "split":
            money = _money(abs(amount), style, "debit" if amount < 0 else "credit")
        elif style.amount_style == "split_zero":
            zero = _money(0, style, "credit")
            debit_text = _money(abs(amount), style, "debit") if amount < 0 else zero
            credit_text = _money(amount, style, "credit") if amount > 0 else zero
            money = f"{debit_text}   {credit_text}"
        else:
            money = _money(abs(amount), style, "debit" if amount < 0 else "credit")
        parts = [date_text, desc] + ([money] if amount else [])
        entry = [parts]  # the entry's lines, kept together on one page
        if fee is not None:
            if style.fee_line:
                parts = ["SERVICE FEE"]
                entry.append(parts)
            parts.append(_money(fee, style, "debit"))
        if show_balance:
            parts.append(_money(balance, style, "balance"))
            if accrued is not None:
                parts.append(_fmt_number(accrued, style.thousands, style.decimal))
        table.append("\n".join("   ".join(p) for p in entry))
        amount_balance = None if fee is not None else (balance if show_balance else None)
        rows_before = len(truth)
        if amount:
            truth.append(TruthRow(d.isoformat(), -amount if amount < 0 else None, amount if amount > 0 else None,
                                  amount_balance))
        if fee is not None:
            truth.append(TruthRow(d.isoformat(), fee, None, balance if show_balance else None))
        rows_per_line.append(len(truth) - rows_before)

    period = f"{_fmt_date(recipe.start, 'dd Mon yyyy')} to {_fmt_date(recipe.end, 'dd Mon yyyy')}"
    head = [recipe.bank, f"{recipe.account_label} {recipe.account}",
            period if style.fee_line else f"Statement Period: {period}"]
    head += recipe.summary_lines
    head.append(f"{style.opening_label}   {_money(recipe.opening, style, 'balance')}")
    header_row = "Date   Description   " + ("Debit   Credit   " if style.amount_style.startswith("split") else
                                           "Amount   ") + ("Fees   " if style.fee_column else "") + "Balance"

    pages: list[list[str]] = []
    rows_left = list(table)
    lines_done = 0
    while rows_left or not pages:
        page = (head + [header_row]) if not pages else [f"{recipe.bank} - continued", header_row]
        if pages and lines_done:
            truth_done = sum(rows_per_line[:lines_done])
            last_printed = next((t.balance for t in reversed(truth[:truth_done]) if t.balance is not None), None)
            if last_printed is not None:
                page.append(f"Balance brought forward   {_money(last_printed, style, 'balance')}")
        chunk, rows_left = rows_left[:style.rows_per_page], rows_left[style.rows_per_page:]
        page += [line for entry in chunk for line in entry.split("\n")]
        lines_done += len(chunk)
        pages.append(page)
    closing = balance
    if style.closing_label:
        pages[-1].append(f"{style.closing_label}   {_money(closing, style, 'balance')}")
    if recipe.totals:
        fees = round(sum(fee or 0 for _, _, _, fee, _ in day_rows), 2) if style.fee_line else 0
        debits = sum(t.debit or 0 for t in truth) - fees  # fee lines: fees are totalled on their own
        credits = sum(t.credit or 0 for t in truth)
        pages[-1].append(f"Total debits   {_fmt_number(debits, style.thousands, style.decimal)}")
        pages[-1].append(f"Total credits   {_fmt_number(credits, style.thousands, style.decimal)}")
        if style.fee_line:
            vat = _fmt_number(fees * 15 / 115, "", style.decimal)
            pages[-1].append(f"Total service fees (R{vat} VAT included)   "
                             f"{_fmt_number(fees, style.thousands, style.decimal)}")
    pages[-1].append("Please report any errors within 30 days.")
    statement = Statement(style, pages, truth, recipe.account, recipe.opening, closing, table=table,
                          truth_rows_per_line=rows_per_line)
    statement.recipe = recipe
    return statement


def render_pdf(statement: Statement, path: Path) -> Path:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas

    c = canvas.Canvas(str(path), pagesize=A4)
    for page in statement.lines_by_page:
        y = 810
        for line in page:
            c.setFont("Helvetica", 7.5)
            c.drawString(28, y, line)
            y -= 11.5
        c.showPage()
    c.save()
    return path


def render_scan(pdf_path: Path, path: Path, rng: random.Random, upside_down: bool = False) -> Path:
    """An image-only copy of the PDF, like a scan: rasterised, noisy, maybe
    upside down."""
    import pymupdf
    from PIL import Image, ImageFilter

    src = pymupdf.open(pdf_path)
    out = pymupdf.open()
    for page in src:
        pix = page.get_pixmap(dpi=200, colorspace=pymupdf.csGRAY)
        img = Image.frombytes("L", (pix.width, pix.height), pix.samples)
        img = img.rotate(rng.uniform(-0.6, 0.6), fillcolor=255, expand=False)
        noise = Image.effect_noise(img.size, 18).point(lambda v: 0 if v < 40 else 255)
        img = Image.composite(img, noise, noise).filter(ImageFilter.GaussianBlur(0.4))
        if upside_down:
            img = img.rotate(180)
        png = path.with_suffix(f".p{page.number}.png")
        img.save(png)
        new = out.new_page(width=page.rect.width, height=page.rect.height)
        new.insert_image(new.rect, filename=str(png))
        png.unlink()
    out.save(path)
    return path
