"""Layout-independent statement reader.

Works on any bank's statement without knowing its layout, using the one
thing every statement has: each transaction's balance equals the previous
balance plus or minus its amount. For each line that looks like a
transaction, every plausible reading of its numbers is tried - which is the
amount and which the balance, whether there's a fee or an extra column,
whether "3 205.00" is one number - and the reading that continues the
balance chain wins, which also shows whether money went in or out.

Nothing is dropped and nothing is guessed: printed figures are kept exactly as
printed, and a row the arithmetic can't confirm (or that could be read two
ways that both add up) is kept and marked with why it needs review.
"""
from __future__ import annotations

import itertools
import re
from dataclasses import dataclass, field
from datetime import date, timedelta

TOLERANCE = 0.005
MAX_UNKNOWN_SIGNS = 10  # rows without a balance whose direction is worked out together

_MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], start=1)}
_FULL_MONTHS = {"january", "february", "march", "april", "may", "june", "july", "august", "september",
                "october", "november", "december", "sept"}

# --- Money ------------------------------------------------------------------------

# Figures are written "1,234.56" or, with a decimal comma, "1 234,56" /
# "1.234,56"; a statement is read in the style most of its figures use
# (decimal_mark), so "12,50" is never read as twelve thousand-odd.
def _money_re(decimal: str) -> re.Pattern:
    num = (r"\d{1,3}(?:\.\d{3})+,\d{2}|\d+,\d{2}" if decimal == ","
           else r"\d{1,3}(?:,\d{3})+\.\d{2}|\d+\.\d{2}")
    return re.compile(
        r"(?<![\w.,/:])"
        r"(?P<pre>\(|-\s?)?(?:R\s?|ZAR\s?)?(?P<pre2>-)?"
        rf"(?P<num>{num})"
        rf"(?P<post>\)|-(?![\d{re.escape(decimal)}])|\s?(?:CR|DR|Cr|Dr|cr|dr)(?![A-Za-z]))?"
        r"\*?"
        r"(?![\w.,/%])"
    )


_MONEY_RES = {".": _money_re("."), ",": _money_re(",")}
# Thousands written out with the separator that isn't the decimal mark
# ("1,234.56", "1.234,56"): where a statement does that, "1 321.75" can't be one number.
_GROUPED_THOUSANDS = {".": re.compile(r"(?<![\d,])\d{1,3}(?:,\d{3})+\.\d{2}(?!\d)"),
                      ",": re.compile(r"(?<![\d.])\d{1,3}(?:\.\d{3})+,\d{2}(?!\d)")}
# "3 205.00" printed with a space as thousands separator: the digits before it.
_SPACE_GROUPS = re.compile(r"(?:^|(?<=[\s:(]))(?P<neg>-\s?|\()?(?:R\s?)?(?P<groups>\d{1,3}(?: \d{3})*) $")


def decimal_mark(text: str) -> str:
    """"," when most of the statement's figures have a decimal comma, else "."."""
    commas = sum(1 for _ in _MONEY_RES[","].finditer(text or ""))
    return "," if commas > sum(1 for _ in _MONEY_RES["."].finditer(text or "")) else "."


@dataclass
class Money:
    value: float  # size, always positive
    sign: int | None  # +1 / -1 when the statement marks it (Cr/Dr, -, brackets), else None
    start: int
    end: int
    # This reading ignores digits just before it that could be its thousands
    # ("3" in "3 250.00"); used only to break ties between readings that add up.
    leaves_digits: bool = False
    # The statement marks credits "Cr" and never debits, so by its own
    # convention an unmarked amount here is money out.
    unmarked_is_out: bool = False
    # The statement marks money out (-, brackets, trailing minus) and never
    # marks money in, so by its own convention an unmarked amount is money in.
    unmarked_is_in: bool = False

    def direction_options(self) -> list[int]:
        if self.sign is not None:
            return [self.sign]
        if self.unmarked_is_out:
            return [-1]
        if self.unmarked_is_in:
            return [1]
        return [1, -1]


def money_tokens(line: str, decimal: str = ".") -> list[list[Money]]:
    """Money on the line, each as its possible readings (a second one when
    "3 205.00" could be one number)."""
    tokens = []
    for m in _MONEY_RES[decimal].finditer(line):
        num = m.group("num")
        whole, cents = num.rsplit(decimal, 1)
        value = float(re.sub(r"\D", "", whole) + "." + cents)
        pre = (m.group("pre") or "") + (m.group("pre2") or "")
        post = (m.group("post") or "").strip().lower()
        sign = None
        if "(" in pre or "-" in pre or post in (")", "-", "dr"):
            sign = -1
        elif post == "cr":
            sign = 1
        readings = [Money(value, sign, m.start(), m.end())]
        if not pre and len(whole) == 3:
            g = _SPACE_GROUPS.search(line[:m.start()])
            if g:
                merged = float(g.group("groups").replace(" ", "") + whole + "." + cents)
                readings[0].leaves_digits = True
                readings.append(Money(merged, -1 if g.group("neg") else sign, g.start(), m.end()))
        tokens.append(readings)
    return tokens


# --- Dates ------------------------------------------------------------------------

_MON = r"[A-Za-z]{3,9}"
_DATE_KINDS = {
    "dmy": r"(?P<d>\d{1,2})[/.-](?P<m>\d{1,2})[/.-](?P<y>\d{4}|\d{2})(?![\d/.-])",
    "ymd": r"(?P<y>\d{4})[/.-](?P<m>\d{1,2})[/.-](?P<d>\d{1,2})(?!\d)",
    "d_mon_y": rf"(?P<d>\d{{1,2}})[ -]?(?P<mon>{_MON})\.?[ ,-]\s?(?P<y>\d{{4}}|\d{{2}})(?![\d.,:])",
    "d_mon": rf"(?P<d>\d{{1,2}})[ -]?(?P<mon>{_MON})(?![A-Za-z])",
    "mon_d_y": rf"(?P<mon>{_MON})\.?\s(?P<d>\d{{1,2}}),?\s(?P<y>\d{{4}})(?!\d)",
    "mon_d": rf"(?P<mon>{_MON})\.?\s(?P<d>\d{{1,2}})(?![\d.,:])",
    "dm": r"(?P<d>\d{1,2})/(?P<m>\d{1,2})(?![/\d])",
}
_DATE_AT_START = {k: re.compile(r"^\W{0,3}" + f"(?:{p})(?=\\s|$)") for k, p in _DATE_KINDS.items()}
_KIND_ORDER = ["dmy", "ymd", "d_mon_y", "mon_d_y", "d_mon", "mon_d", "dm"]


def _month_number(text: str) -> int | str | None:
    """1-12; "near" for an OCR-damaged month ("Fbe", "Der"); None if not a month."""
    word = text.lower().rstrip(".")
    if word[:3] in _MONTHS and (len(word) <= 4 or word in _FULL_MONTHS):
        return _MONTHS[word[:3]]
    if len(word) == 3 and any(sum(a != b for a, b in zip(word, m)) <= 1 for m in _MONTHS):
        return "near"
    return None


@dataclass
class DateHit:
    end: int
    raw: str
    day: int | None = None  # None: a date is there but unreadable ("41 Mar", "10 Der")
    month: int | None = None
    year: int | None = None


def _match_date(line: str, kind: str, month_first: bool) -> DateHit | None:
    m = _DATE_AT_START[kind].match(line)
    if not m:
        return None
    g = m.groupdict()
    raw = m.group(0).strip(" -|([")
    if g.get("mon") is not None:
        month = _month_number(g["mon"])
        if month is None:
            return None
        if month == "near":
            return DateHit(m.end(), raw)
    else:
        month = int(g["m"])
    day = int(g["d"])
    if month_first and g.get("m") is not None:
        day, month = month, day
    year = int(g["y"]) if g.get("y") else None
    if year is not None and year < 100:
        year += 2000
    try:
        date(year or 2000, month, day)  # 2000 is a leap year, so 29 Feb passes without a year
    except ValueError:
        return DateHit(m.end(), raw)
    return DateHit(m.end(), raw, day, month, year)


def _choose_date_kind(lines: list[str], decimal: str = ".") -> tuple[str | None, bool]:
    """The date style the statement's rows use, and whether it's month-first."""
    counts: dict[str, int] = {}
    day_first = month_first = 0
    for line in lines:
        if not _MONEY_RES[decimal].search(line):
            continue
        for kind in _KIND_ORDER:
            hit = _match_date(line, kind, False)
            if not hit:
                continue
            swapped = _match_date(line, kind, True) if kind in ("dmy", "dm") and hit.day is None else None
            if swapped and swapped.day is not None:
                month_first += 1  # 03/25/2026 only reads month-first
            elif kind in ("dmy", "dm") and hit.day and hit.day > 12:
                day_first += 1
            if hit.day is not None or kind in ("dmy", "dm"):
                counts[kind] = counts.get(kind, 0) + 1
                break
    if not counts:
        return None, False
    kind = max(counts, key=lambda k: (counts[k], -_KIND_ORDER.index(k)))
    return kind, month_first > day_first


def _latest_full_date(lines: list[str]) -> date | None:
    found = []
    for line in lines:
        for kind in ("dmy", "ymd", "d_mon_y", "mon_d_y"):
            for m in re.finditer(_DATE_KINDS[kind], line):
                hit = _match_date(m.group(0), kind, False)
                if hit and hit.day and hit.year:
                    found.append(date(hit.year, hit.month, hit.day))
    plausible = [d for d in found if 2000 <= d.year <= date.today().year + 1]
    return max(plausible) if plausible else None


# --- Result -----------------------------------------------------------------------------

@dataclass
class AutoRow:
    date_iso: str | None  # None when no readable date is printed on the line
    description: str
    debit: float | None
    credit: float | None
    balance: float | None  # as printed
    check: str = ""  # why this row needs review; "" when the balances confirm it
    evidence: str = ""  # the line exactly as read from the document
    # An amount whose direction (in or out) the document doesn't establish:
    # kept here rather than guessed into debit or credit.
    unassigned: float | None = None


@dataclass
class AutoResult:
    rows: list[AutoRow]
    opening_balance: float | None
    closing_balance: float | None


_OPENING_WORDS = re.compile(
    r"opening balance|balance brought forward|brought forward|balance b/?f\b|b/fwd|balance forward|"
    r"previous balance|balance at start|starting balance|start balance", re.IGNORECASE)
_CLOSING_WORDS = re.compile(
    r"closing balance|balance carried forward|carried forward|balance c/?f\b|c/fwd|balance at end|"
    r"ending balance|end balance|new balance", re.IGNORECASE)
_NOT_A_ROW = re.compile(r"\b(totals?|turnover|summary|available balance|interest rate)\b", re.IGNORECASE)
_NOT_DESCRIPTION = re.compile(r"\b(page|date|description|balance|statement|continued|turn over)\b", re.IGNORECASE)
_DIGIT = re.compile(r"\d")

UNKNOWN_DIRECTION = 0


@dataclass
class _Line:
    index: int
    text: str
    date: DateHit | None
    tokens: list[list[Money]]  # money after the date
    # The dated line whose entry this undated line finishes; its date is this line's.
    continues: "_Line | None" = None


@dataclass
class _Reading:
    amount: Money
    fee: Money | None
    balance: Money
    trailing: int  # an extra column after the balance (e.g. accrued charges)


def _readings(line: str, tokens: list[list[Money]]) -> list[_Reading]:
    """Ways of reading a row's money, simplest first: amount then balance;
    amount, fee, balance; and either with an extra column after the balance.
    A reading that leaves digits unexplained between the amount, fee and
    balance columns is not a reading of what's printed, so it's left out."""
    out: list[_Reading] = []
    # "0.00" printed in an empty debit or credit column is a blank column.
    zeros = [t for t in tokens[:-1] if all(r.value == 0 for r in t)]
    blanked = list(line)
    for t in zeros:
        blanked[t[0].start:t[0].end] = " " * (t[0].end - t[0].start)
    blanked = "".join(blanked)

    def clean_gap(a: Money, b: Money) -> bool:
        return a.end <= b.start and not _DIGIT.search(blanked[a.end:b.start])

    seen = set()
    for toks in ([tokens, [t for t in tokens if t not in zeros]] if zeros else [tokens]):
        k = len(toks)
        for trailing in (0, 1):
            b = k - 1 - trailing
            if b < 1:
                continue
            with_fee = []
            for choice in itertools.product(*toks[max(0, b - 2):b + 1]):
                bal, amt = choice[-1], choice[-2]
                if clean_gap(amt, bal):
                    out.append(_Reading(amt, None, bal, trailing))
                if len(choice) == 3 and clean_gap(choice[0], amt) and clean_gap(amt, bal):
                    with_fee.append(_Reading(choice[0], amt, bal, trailing))
            out.extend(with_fee)
    unique = []
    has_amount = any(any(r.value for r in t) for t in tokens[:-1])
    for r in out:
        if r.amount.value == 0 and (r.fee is None or r.fee.value == 0) and has_amount:
            continue  # "0.00" is an empty column here, not the amount
        key = (r.amount.start, r.amount.value, r.fee and (r.fee.start, r.fee.value), r.balance.start,
               r.balance.value, r.trailing)
        if key not in seen:
            seen.add(key)
            unique.append(r)
    return unique


def _balance_values(money: Money, allow_flip: bool) -> list[float]:
    if money.sign is not None:
        return [money.sign * money.value]
    return [money.value, -money.value] if allow_flip else [money.value]


def _fits(prev: float, reading: _Reading, balance: float) -> int | None:
    """Direction of the amount (+1 in, -1 out) if the reading continues the chain."""
    fee = 0.0
    if reading.fee is not None:
        fee = reading.fee.value if reading.fee.sign == 1 else -reading.fee.value
    for direction in reading.amount.direction_options():
        if abs(prev + direction * reading.amount.value + fee - balance) <= TOLERANCE:
            return direction
    return None


BEAM_WIDTH = 6  # competing readings of the statement kept at once


@dataclass
class _Path:
    """One consistent way of reading the statement so far."""
    state: float | None = None  # balance after the last row, per the arithmetic
    opening: float | None = None
    # (line, readings, kind): amounts waiting for a balance to confirm them.
    # kind "amount": a dated line with one amount; "listed": the same, under a
    # list heading (upcoming debit orders, pending...); "row": a line whose
    # figures didn't fit as amount + balance, tried as balance-less amounts.
    pending: list = field(default_factory=list)
    ops: list = field(default_factory=list)  # rows, built at the end
    flags: int = 0
    stray_digits: int = 0  # tie-breaker: readings that left a digit group before the amount unexplained

    def fork(self, **changes) -> "_Path":
        new = _Path(self.state, self.opening, list(self.pending), list(self.ops), self.flags, self.stray_digits)
        for key, value in changes.items():
            setattr(new, key, value)
        return new

    def add(self, line, direction, amount, balance, upto, check="", prefix=""):
        self.ops.append((line, direction, amount, balance, upto, check, prefix))
        if check:
            self.flags += 1

    def add_reading(self, line, reading, direction, balance, check=""):
        """balance: as printed on the row."""
        fee = reading.fee
        if fee is None or fee.value == 0:
            # A zero amount moves no money (e.g. a rate notice) - unless the row
            # needs review, which must never disappear.
            if reading.amount.value or check:
                self.add(line, direction, reading.amount.value, balance, reading.amount.start, check)
            return
        fee_direction = 1 if fee.sign == 1 else -1
        if reading.amount.value or check:
            # The balance between the amount and the fee isn't printed.
            self.add(line, direction, reading.amount.value, None, reading.amount.start, check)
        self.add(line, fee_direction, fee.value, balance, reading.amount.start, check, prefix="Fee: ")

    def set_aside(self, entries) -> None:
        """Pending lines the balances say did not move money. Lines under a
        list heading are a list, not transactions; any other line is kept -
        never silently dropped - and marked for review."""
        seen = set()
        for line, readings, kind in entries:
            if kind == "listed" or id(line) in seen:
                continue
            seen.add(id(line))
            money = line.tokens[0][0]
            self.add(line, UNKNOWN_DIRECTION, money.value, None, money.start,
                     "these figures don't fit the running balance (a duplicate, misprint or misread line?) "
                     "- not counted in any total")

    def flush_pending(self, reason):
        """Rows printed without a balance that no later balance confirmed:
        kept as printed, never approved. Lines under a list heading are a
        list, not transactions."""
        for line, readings, kind in self.pending:
            if kind == "listed":
                continue
            money = readings[0]
            known = money.sign is not None
            self.add(line, money.sign if known else UNKNOWN_DIRECTION, money.value, None, money.start,
                     f"no balance confirms this amount ({reason})" if known else
                     f"the statement doesn't show whether this money came in or went out, and no balance "
                     f"confirms it ({reason})")
        self.pending = []


def extract(text: str, period_end: date | None = None, period_start: date | None = None) -> AutoResult:
    lines = [ln.strip() for ln in (text or "").splitlines()]
    decimal = decimal_mark(text)
    kind, month_first = _choose_date_kind(lines, decimal)
    reference_end = period_end or _latest_full_date(lines) or date.today()
    reference_start = period_start if period_end else None

    parsed: list[_Line] = []
    for i, line in enumerate(lines):
        hit = _match_date(line, kind, month_first) if kind else None
        if hit and hit.year and abs(hit.year - reference_end.year) > 1 and kind == "d_mon_y":
            hit = _match_date(line, "d_mon", False)  # "01 Sep 20 litres": 20 is not a year here
        start = hit.end if hit else 0
        parsed.append(_Line(i, line, hit, [t for t in money_tokens(line, decimal) if t[0].start >= start]))

    if _GROUPED_THOUSANDS[decimal].search(text or ""):
        # The statement writes thousands with commas (or, with a decimal
        # comma, dots), so "1 321.75" can't be one number here: drop those readings.
        for p in parsed:
            p.tokens = [[t[0]] for t in p.tokens]
            for t in p.tokens:
                t[0].leaves_digits = False
    _link_entry_lines(parsed)
    dated = [p for p in parsed if p.date and p.tokens]
    # Statements that mark credits "Cr" print overdrawn balances unmarked, so
    # an unmarked balance may be negative there.
    allow_flip = any(r.sign == 1 for p in dated for tok in p.tokens for r in tok[:1])
    if allow_flip and not any(r.sign == -1 and "dr" in p.text[r.start:r.end].lower()
                              for p in dated for tok in p.tokens for r in tok[:1]):
        for p in dated:
            for tok in p.tokens:
                for r in tok:
                    r.unmarked_is_out = True
    if not allow_flip:
        # Amounts (not the balance at the end of the line) marked as money out,
        # and none marked as money in: unmarked amounts are money in.
        movements = [r for p in dated for tok in p.tokens[:-1] for r in tok[:1]]
        if any(r.sign == -1 for r in movements) and not any(r.sign == 1 for r in movements):
            for p in dated:
                for tok in p.tokens:
                    for r in tok:
                        r.unmarked_is_in = True
    # Each a list of possible values ("R3 304.56" could be 304.56 or 3,304.56)
    # with the digits each reading leaves unexplained.
    openings = _find_balance_readings(lines[:(dated[0].index + 1) if dated else len(lines)], _OPENING_WORDS, True,
                                      allow_flip, decimal)
    closing_readings = _find_balance_readings(lines, _CLOSING_WORDS, False, allow_flip, decimal)
    closings = [value for value, _ in closing_readings]
    if not dated:
        return AutoResult([], openings[0][0] if openings else None, closings[0] if closings else None)
    first, last = dated[0].index, dated[-1].index
    if not openings:
        line = next((p for p in parsed[first:last + 1] if p.tokens and _OPENING_WORDS.search(p.text)), None)
        openings = ([(v, r.leaves_digits) for r in line.tokens[-1] for v in _balance_values(r, allow_flip)]
                    if line else [])

    listed = _listed_lines(parsed)
    paths = [_Path()]
    for p in parsed[first:last + 1]:
        if not p.tokens or _OPENING_WORDS.search(p.text) or _CLOSING_WORDS.search(p.text):
            continue
        if not p.date and _NOT_A_ROW.search(p.text):
            continue
        nxt = next((q for q in parsed[p.index + 1:last + 1] if len(q.tokens) >= 2 and q.date), None)
        stepped = []
        for path in paths:
            stepped.extend(_step(path, p, openings, nxt, allow_flip, listed))
        paths = _prune(stepped)

    finished = []
    for path in paths:
        path = path.fork()
        mismatch = 0
        if path.pending:
            if any(abs(c - (path.state or 0)) <= TOLERANCE for c in closings):
                # Already at the closing balance: these lines didn't move money.
                path.set_aside(path.pending)
                path.pending = []
            else:
                path.flush_pending("no balance printed after it")
        strays = path.stray_digits
        if closing_readings and path.state is not None:
            matched = [left for c, left in closing_readings if abs(c - path.state) <= TOLERANCE]
            if matched:
                strays += min(matched)  # the closing balance read by dropping a digit group
            else:
                mismatch = 1
        finished.append(((path.flags + strays + mismatch, strays), path))
    finished.sort(key=lambda f: f[0])
    best = finished[0][1]

    rows = _build_rows(best.ops, lines, kind, month_first, reference_end, reference_start, decimal)
    _flag_ambiguity(rows, best, [p for score, p in finished[1:] if score[0] == finished[0][0][0]])
    closing = next((c for c in closings if best.state is not None and abs(c - best.state) <= TOLERANCE),
                   closings[0] if closings else None)
    opening = best.opening if best.opening is not None else (openings[0][0] if openings else None)
    return AutoResult(rows, opening, closing)


def _link_entry_lines(parsed: list[_Line]) -> None:
    """An entry can run over two lines with the date printed once: the dated
    line prints no balance, and the next line finishes it with a figure and
    the balance ("3 Mar Payment 1 200,00-" then "Service Fee 8,75- 2 291,25";
    "4 Mar Payment notification" then "Service Fee 1,25- 2 290,00"). That
    next line is part of the dated entry, so it has the entry's date."""
    previous = None
    for p in parsed:
        if not p.text:
            continue
        if (p.date is None and len(p.tokens) >= 2 and previous is not None and previous.date is not None
                and previous.continues is None and len(previous.tokens) <= 1
                and not any(w.search(t) for w in (_OPENING_WORDS, _CLOSING_WORDS) for t in (p.text, previous.text))
                and not _NOT_A_ROW.search(p.text)):
            p.date, p.continues = previous.date, previous
        previous = p


_LIST_HEADING = re.compile(
    r"pending|scheduled|upcoming|future|subscriptions?|debit orders|standing orders|recurring|"
    r"not yet (?:processed|cleared)|authoris", re.IGNORECASE)


def _listed_lines(parsed: list[_Line]) -> set[int]:
    """Dated one-amount lines that sit under a list heading ("Upcoming debit
    orders", "Pending card transactions"...) rather than in the transactions."""
    listed = set()
    for i, p in enumerate(parsed):
        if not (p.date and len(p.tokens) == 1):
            continue
        j = i - 1
        while j >= 0 and parsed[j].date and len(parsed[j].tokens) == 1:
            j -= 1
        while j >= 0 and not parsed[j].text:
            j -= 1
        if j >= 0 and not parsed[j].tokens and _LIST_HEADING.search(parsed[j].text):
            listed.add(p.index)
    return listed


def _flag_ambiguity(rows: list[AutoRow], best: _Path, rivals: list[_Path]) -> None:
    """Another reading that adds up just as well but gives different figures:
    the rows where they differ can't be settled from the document."""
    def figures(path):
        return {(op[0].index, op[6]): (op[1], op[2]) for op in path.ops}

    mine = figures(best)
    keys = [(op[0].index, op[6]) for op in best.ops]
    for rival in rivals:
        theirs = figures(rival)
        for i, key in enumerate(keys):
            if key in theirs and theirs[key] != mine[key] and not rows[i].check:
                direction, amount = theirs[key]
                rows[i].check = f"ambiguous: the document also reads as {amount:,.2f} here and still adds up"


def _prune(paths: list[_Path]) -> list[_Path]:
    best_by_state: dict = {}
    for path in paths:
        key = (None if path.state is None else round(path.state, 2), len(path.pending))
        if key not in best_by_state or _cost(path) < _cost(best_by_state[key]):
            best_by_state[key] = path
    return sorted(best_by_state.values(), key=_cost)[:BEAM_WIDTH]


def _cost(path: _Path) -> tuple:
    # A reading that leaves a digit group unexplained ("1" in "1 321.75")
    # weighs as much as a row that doesn't fit; ties are then flagged as
    # ambiguous rather than silently decided.
    return (path.flags + path.stray_digits, path.stray_digits)


def _step(path: _Path, p: _Line, openings: list[tuple[float, int]], nxt: _Line | None, allow_flip: bool,
          listed: set[int]) -> list[_Path]:
    """The ways this line can continue the path."""
    if len(p.tokens) == 1:
        kind = "listed" if p.index in listed else "amount"
        return [path.fork(pending=path.pending + [(p, p.tokens[0], kind)])] if p.date else [path]

    readings = _readings(p.text, p.tokens)
    if not readings:
        return [path] if not p.date else [_unreadable(path, p)]
    if path.state is None and openings:
        out = []
        for opening, left in openings:
            start = path.fork(state=opening, opening=opening)
            start.stray_digits += left  # read by leaving a digit group of the printed figure unexplained
            out.extend(_step(start, p, [], nxt, allow_flip, listed))
        return out
    if path.state is None:
        # No opening balance printed: the chain starts at this row, and the
        # first amount's direction can't be established from the balances.
        new = path.fork()
        reading = readings[0]
        balance = _balance_values(reading.balance, False)[0]
        new.flush_pending("no opening balance printed")
        known = reading.amount.sign is not None
        new.add_reading(p, reading, reading.amount.sign if known else UNKNOWN_DIRECTION, balance,
                        "" if known else "no opening balance is printed, so this first amount can't be "
                                         "confirmed as money in or out")
        new.state = balance
        return [new]

    options = _fitting(path.state, path.pending, readings, allow_flip)
    pending, set_aside = path.pending, []
    if not options and pending:
        # The chain works without the amount-only lines before this row, so
        # they did not move money.
        options = _fitting(path.state, [], readings, allow_flip)
        pending, set_aside = [], path.pending
    if options:
        out = []
        for signs, ambiguous, reading, direction, balance in options:
            new = path.fork(pending=[])
            new.set_aside(set_aside)
            for (line, _, _), (money, sign) in zip(pending, signs):
                if ambiguous:
                    # More than one way to read these rows adds up to the balance.
                    new.add(line, money.sign if money.sign else UNKNOWN_DIRECTION, money.value, None, money.start,
                            "more than one reading of the rows before the next balance adds up - check them")
                else:
                    new.add(line, sign, money.value, None, money.start)
                    new.stray_digits += money.leaves_digits
            new.add_reading(p, reading, direction, balance,
                            "more than one reading of this row adds up - check it" if ambiguous else "")
            new.stray_digits += _left_over(reading)
            new.state = balance
            out.append(new)
        return out

    if not p.date:
        return [path]  # money on a line that is not a transaction: notes, summaries, carry-overs

    as_pending = path.fork(pending=path.pending + [(p, tok, "row") for tok in p.tokens if any(r.value for r in tok)])
    new = path.fork()
    new.flush_pending("a later row does not add up")
    # Nothing confirms a reading here, so the row is kept as printed: whole
    # figures first, not ones that leave a digit group behind.
    readings = sorted(readings, key=_left_over)
    printed = _balance_values(readings[0].balance, False)[0]
    found = _explain(new.state, readings, nxt, allow_flip)
    if found is not None:
        reading, direction, printed, state, what = found
    else:
        reading = readings[0]
        direction = reading.amount.sign if reading.amount.sign is not None else UNKNOWN_DIRECTION
        state = printed
        what = "does not follow from the previous balance - a row may be missing or misread"
    # The row is kept exactly as printed; the review note says what doesn't fit.
    new.add_reading(p, reading, direction, printed, what)
    new.state = state
    return [new, as_pending]


def _left_over(reading: _Reading) -> int:
    return sum(m.leaves_digits for m in (reading.amount, reading.fee, reading.balance) if m)


def _unreadable(path: _Path, p: _Line) -> _Path:
    new = path.fork()
    money = p.tokens[-1][0]
    new.add(p, UNKNOWN_DIRECTION, money.value, None, money.start,
            "the amounts on this line could not be told apart - check it against the document")
    return new


MAX_PENDING_COMBINATIONS = 4096


def _fitting(state, pending, readings, allow_flip):
    """Every (pending choices, ambiguous, reading, direction, balance) that
    continues the chain - one per resulting balance. Rows printed without a
    balance (pending) are confirmed together by the next balance: each can
    be any of its readings, in or out unless marked."""
    options = []
    for _, monies, _ in pending:
        options.append([(m, s) for m in monies for s in m.direction_options()])
    combos = 1
    for o in options:
        combos *= len(o)
    if combos > MAX_PENDING_COMBINATIONS:
        return []
    by_balance: dict = {}  # balance -> [(stray digits, choice, reading, direction)]
    for reading in readings:
        reading_strays = sum(m.leaves_digits for m in (reading.amount, reading.fee, reading.balance) if m)
        for balance in _balance_values(reading.balance, allow_flip):
            for choice in itertools.product(*options):
                start = state + sum(sign * money.value for money, sign in choice)
                direction = _fits(start, reading, balance)
                if direction is not None:
                    strays = reading_strays + sum(m.leaves_digits for m, _ in choice)
                    by_balance.setdefault(round(balance, 2), []).append((strays, choice, reading, direction, balance))
    out = []
    for candidates in by_balance.values():
        # Of the readings that reach this balance, the one leaving the fewest
        # digits unexplained; ambiguous if another is just as good but differs.
        fewest = min(c[0] for c in candidates)
        best = [c for c in candidates if c[0] == fewest]
        _, choice, reading, direction, balance = best[0]
        out.append((choice, len({_movements(c) for c in best}) > 1, reading, direction, balance))
    return out


def _movements(candidate) -> tuple:
    """The money a candidate reading moves (zeros left out), to tell readings
    that differ from ones that only split the same money differently."""
    _, choice, reading, direction, _ = candidate
    moves = [(sign, round(m.value, 2)) for m, sign in choice if m.value]
    if reading.amount.value:
        moves.append((direction, round(reading.amount.value, 2)))
    if reading.fee is not None and reading.fee.value:
        moves.append((1 if reading.fee.sign == 1 else -1, round(reading.fee.value, 2)))
    return tuple(sorted(moves))


def _explain(state, readings, nxt: _Line | None, allow_flip):
    """This row doesn't continue the chain but the next row does, from one of
    two balances. Works out which printed figure disagrees, without changing
    it: (reading, direction, printed balance, balance to continue from, note)."""
    if nxt is None:
        return None
    next_readings = _readings(nxt.text, nxt.tokens)

    def next_fits(balance: float) -> bool:
        return any(_fits(balance, r, b) is not None for r in next_readings
                   for b in _balance_values(r.balance, allow_flip))

    for reading in readings:
        size = reading.amount.value + (reading.fee.value if reading.fee else 0)
        printed = _balance_values(reading.balance, False)[0]
        for direction in reading.amount.direction_options():
            expected = round(state + direction * size, 2)
            if next_fits(expected):
                return (reading, direction, printed, expected,
                        f"balance printed as {printed:,.2f} but the amounts give {expected:,.2f} - "
                        f"possible misprint or misread")
        for balance in _balance_values(reading.balance, allow_flip):
            if next_fits(balance):
                delta = round(balance - state, 2)
                if abs(delta) <= TOLERANCE:
                    continue
                direction = reading.amount.sign if reading.amount.sign is not None else UNKNOWN_DIRECTION
                return (reading, direction, balance, balance,
                        f"amount printed as {size:,.2f} but the balances give {abs(delta):,.2f} - "
                        f"possible misprint or misread")
    return None


def _build_rows(ops, lines, kind, month_first, reference_end, reference_start=None, decimal=".") -> list[AutoRow]:
    rows = []
    amount_at: dict[int, int] = {}  # where each line's amount was read to start
    for op in ops:
        amount_at.setdefault(op[0].index, op[4])
    for line, direction, amount, balance, upto, check, prefix in ops:
        date_check = ""
        if line.date is None:
            iso, date_check = None, "no date printed on this line"
        elif line.date.day is None:
            iso, date_check = None, f"date unreadable ({line.date.raw})"
        else:
            iso = _with_year(line.date, reference_end, reference_start)
        entry = line.continues
        desc = line.text[(line.date.end if line.date and not entry else 0):upto].strip(" |")
        if entry:
            # The entry's own words, then this line's ("Proof of payment SMS | Service Fee").
            end = amount_at.get(entry.index)
            if end is None and entry.tokens:
                end = min(r.start for r in entry.tokens[0])
            words = entry.text[entry.date.end:end].strip(" |")
            desc = " | ".join(part for part in (words, desc) if part)
        for extra in lines[line.index + 1: line.index + 3]:
            if (not extra or len(extra) > 70 or money_tokens(extra, decimal)
                    or (kind and _match_date(extra, kind, month_first))
                    or _NOT_DESCRIPTION.search(extra) or extra.startswith("*")):
                break
            desc = f"{desc} | {extra}" if desc else extra
        rows.append(AutoRow(
            iso, prefix + (desc or "(no description)"),
            amount if direction < 0 else None, amount if direction > 0 else None, balance,
            "; ".join(c for c in (check, date_check) if c), f"{entry.text} / {line.text}" if entry else line.text,
            amount if direction == UNKNOWN_DIRECTION else None,
        ))
    return rows


def _with_year(hit: DateHit, reference_end: date, reference_start: date | None = None) -> str | None:
    """A date without a year gets the year from the statement period it's in."""
    if hit.year:
        return date(hit.year, hit.month, hit.day).isoformat()
    found = date_in_period(hit.month, hit.day, reference_start, reference_end)
    return found.isoformat() if found else None


def date_in_period(month: int, day: int, period_start: date | None, period_end: date) -> date | None:
    """The date a year-less "11 Aug" falls on. With the statement period
    known: the year that puts it inside the period, or else the one nearest
    to it (a row dated a day after the period ends is still that year) - so a
    6- or 12-month statement crossing New Year gets each month's own year.
    With only the end known: the end's year, or the year before for dates
    more than a month after the end."""
    if period_start is None:
        try:
            candidate = date(period_end.year, month, day)
        except ValueError:  # 29 Feb in a non-leap year: not a real date there
            return None
        if candidate > period_end + timedelta(days=31):
            candidate = candidate.replace(year=candidate.year - 1)
        return candidate
    candidates = []
    for year in (period_end.year, period_end.year - 1, period_end.year - 2):
        try:
            candidates.append(date(year, month, day))
        except ValueError:  # 29 Feb outside a leap year
            continue

    def distance(d: date) -> int:
        if period_start <= d <= period_end:
            return 0
        return min(abs((d - period_start).days), abs((d - period_end).days))

    return min(candidates, key=distance) if candidates else None


def _find_balance(lines: list[str], words: re.Pattern, first: bool, allow_flip: bool = False,
                  decimal: str = ".") -> list[float]:
    """Possible values of the first (or last) balance after the words."""
    return [value for value, _ in _find_balance_readings(lines, words, first, allow_flip, decimal)]


def _find_balance_readings(lines: list[str], words: re.Pattern, first: bool,
                           allow_flip: bool = False, decimal: str = ".") -> list[tuple[float, int]]:
    """As _find_balance, each value with the digit groups its reading leaves
    unexplained ("1 321.75" read as 321.75 leaves the "1")."""
    hits = []
    for line in lines:
        m = words.search(line)
        if not m:
            continue
        tokens = money_tokens(line[m.end():], decimal)
        if tokens:
            hits.append([(v, r.leaves_digits) for r in tokens[0] for v in _balance_values(r, allow_flip)])
    if not hits:
        return []
    return hits[0] if first else hits[-1]


def printed_closings(text: str) -> list[float]:
    """Possible values of the closing balance printed on the statement, read
    with the statement's own sign convention."""
    lines = [ln.strip() for ln in (text or "").splitlines()]
    decimal = decimal_mark(text)
    allow_flip = any(r.sign == 1 for ln in lines for tok in money_tokens(ln, decimal) for r in tok[:1])
    return _find_balance(lines, _CLOSING_WORDS, False, allow_flip, decimal)
