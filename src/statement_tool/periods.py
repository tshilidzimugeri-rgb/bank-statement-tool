"""Reporting periods: months, quarters and financial years (South Africa's
tax year, March to February), and two-monthly VAT periods - built from the
months a workbook's transactions fall in. Used by the upload page's period
filter and the workbook's Period Summary sheet."""
from __future__ import annotations

import calendar
from dataclasses import dataclass

MONTH, QUARTER, YEAR, ALL, VAT_PERIOD = "month", "quarter", "year", "all", "vat"


@dataclass(frozen=True)
class Period:
    label: str
    kind: str
    months: tuple[str, ...]  # "YYYY-MM" keys of the months it covers that have transactions
    length: int  # months the period has in full (3 for a quarter...)

    @property
    def partial(self) -> bool:
        return len(self.months) < self.length


def month_label(month: str) -> str:
    year, m = int(month[:4]), int(month[5:7])
    return f"{calendar.month_abbr[m]} {year}"


def financial_year(month: str) -> int:
    """The calendar year a March-to-February financial year starts in."""
    year, m = int(month[:4]), int(month[5:7])
    return year if m >= 3 else year - 1


def _fy_label(start: int) -> str:
    return f"FY{start}/{(start + 1) % 100:02d}"


def periods(months: list[str]) -> list[Period]:
    """All months, then each financial year and its quarters (Q1 Mar-May,
    Q2 Jun-Aug, Q3 Sep-Nov, Q4 Dec-Feb), then each month - only periods
    with transactions, in date order."""
    months = sorted(set(m for m in months if m))
    if not months:
        return []
    out = [Period(f"All months ({month_label(months[0])} - {month_label(months[-1])})", ALL, tuple(months),
                  len(months))]
    for start in sorted({financial_year(m) for m in months}):
        in_year = tuple(m for m in months if financial_year(m) == start)
        out.append(Period(f"{_fy_label(start)} (Mar {start} - Feb {start + 1})", YEAR, in_year, 12))
        for q in range(4):
            first = 3 + 3 * q  # month number, 3..14 (13, 14 = Jan, Feb next year)
            keys = [f"{start + (first + i - 1) // 12}-{(first + i - 1) % 12 + 1:02d}" for i in range(3)]
            have = tuple(m for m in keys if m in in_year)
            if have:
                span = f"{calendar.month_abbr[int(keys[0][5:])]} - {month_label(keys[-1])}"
                out.append(Period(f"{_fy_label(start)} Q{q + 1} ({span})", QUARTER, have, 3))
    out += [Period(month_label(m), MONTH, (m,), 1) for m in months]
    return out


def vat_periods(months: list[str], category: str) -> list[Period]:
    """SARS VAT periods with transactions: category A two months ending
    January, March..., B ending February, April..., C every month."""
    months = sorted(set(m for m in months if m))
    if category == "C":
        return [Period(month_label(m), VAT_PERIOD, (m,), 1) for m in months]
    ends_odd = category == "A"
    out: dict[str, list[str]] = {}
    for m in months:
        year, mm = int(m[:4]), int(m[5:7])
        # The period's last month: this month if it ends one, else the next.
        if (mm % 2 == 1) == ends_odd:
            end_year, end = year, mm
        else:
            end_year, end = (year + 1, 1) if mm == 12 else (year, mm + 1)
        out.setdefault(f"{end_year}-{end:02d}", []).append(m)
    result = []
    for end in sorted(out):
        year, mm = int(end[:4]), int(end[5:7])
        first_year, first = (year - 1, 12) if mm == 1 else (year, mm - 1)
        label = f"{calendar.month_abbr[first]} {first_year} - {calendar.month_abbr[mm]} {year}"
        result.append(Period(label, VAT_PERIOD, tuple(out[end]), 2))
    return result
