"""Upload page: drop in statement PDFs, get the workbook and reports.

Runs on this computer (run_app.bat) or hosted online (Streamlit Community
Cloud). Online it needs a password (APP_PASSWORD in the app's secrets).

A side menu splits it into pages: Upload statements (PDFs and Excel/CSV
statements), Continue a workbook, Overview, Monthly, VAT, Checks and
Download, with a reporting-period picker (all months, a financial year, a
quarter, a month or a custom range). Each visit works in its own temporary
folder, and each upload makes a new workbook from just the statements
uploaded then - statements from earlier uploads never carry over. To add
statements to a workbook downloaded before, open it on "Continue a
workbook" first.

Run with run_app.bat, or:  .venv\\Scripts\\streamlit run src\\statement_tool\\app.py
"""
from __future__ import annotations

import dataclasses
import hmac
import io
import os
import shutil
import sys
import tempfile
import zipfile
from collections import defaultdict
from pathlib import Path

import pandas as pd
import streamlit as st

# Hosted, this file is run straight from the repo without installing the
# package, so make src/ importable.
_SRC = Path(__file__).resolve().parents[1]
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

# After an update the server can keep the previous version of this package's
# modules loaded while running the new page, which then can't find what it
# imports. Whenever the code on disk has changed, import it afresh.
_CODE_STAMP = max(p.stat().st_mtime_ns for p in (_SRC / "statement_tool").rglob("*.py"))
if "statement_tool" in sys.modules and getattr(sys.modules["statement_tool"], "code_stamp", None) != _CODE_STAMP:
    for _name in [n for n in sys.modules if n == "statement_tool" or n.startswith("statement_tool.")]:
        del sys.modules[_name]
import statement_tool  # noqa: E402

statement_tool.code_stamp = _CODE_STAMP

from statement_tool import config as config_mod  # noqa: E402
from statement_tool.categorize import load_categories
from statement_tool.checks import workbook_gaps
from statement_tool.excel_writer import (
    VAT201_FIELDS,
    VAT_RATE,
    MixedAccountsError,
    WorkbookLockedError,
    append_transactions,
    list_workbooks,
    read_reviews,
    read_transactions,
    report_categories,
    restore_workbook,
    statement_review_row,
    statement_workbook_path,
    vat201_by_month,
    vat_by_month,
)
from statement_tool.extract.document import parse_document
from statement_tool.extract.spreadsheet import SPREADSHEET_TYPES, parse_spreadsheet
from statement_tool.periods import Period, month_label, periods, vat_periods
from statement_tool.store import sha256_of_bytes

st.set_page_config(page_title="Bank Statement Tool", page_icon="📄", layout="wide")

settings = config_mod.load_settings()

# run_app.bat binds to localhost; anything else is treated as online.
HOSTED = st.get_option("server.address") not in ("localhost", "127.0.0.1")


def _app_password() -> str | None:
    try:
        if "APP_PASSWORD" in st.secrets:
            return str(st.secrets["APP_PASSWORD"])
    except Exception:  # no secrets file - normal when running locally
        pass
    return os.environ.get("APP_PASSWORD") or None


expected_password = _app_password()
if HOSTED and not expected_password:
    st.error("This page is online but has no password set, so it won't open. "
             "Add APP_PASSWORD in the app's Secrets settings.")
    st.stop()
if expected_password and not st.session_state.get("signed_in"):
    st.title("Bank Statement Tool")
    entered = st.text_input("Password", type="password")
    if entered and hmac.compare_digest(entered.encode(), expected_password.encode()):
        st.session_state["signed_in"] = True
        st.rerun()
    elif entered:
        st.error("Wrong password.")
    st.stop()

# Each visit gets its own temporary folder: nothing is shared between
# visitors, and nothing from an earlier upload is mixed into a new one.
if "work_dir" not in st.session_state:
    st.session_state["work_dir"] = tempfile.mkdtemp(prefix="statements-")
work_dir = Path(st.session_state["work_dir"])
settings = dataclasses.replace(
    settings,
    output_dir=work_dir / "output",
    uploads_dir=work_dir / "uploads",
    processed_db=work_dir / "processed.db",
)
try:
    categories = load_categories(settings.categories_config)
    layouts, generic = config_mod.load_bank_layouts(settings.banks_config)
    client_rules = config_mod.load_client_rules(settings.clients_config)
except Exception as exc:  # a typo in one of the YAML files
    st.error(f"A settings file in config/ has a mistake, so nothing can be processed until it's fixed:\n\n{exc}")
    st.stop()


# --- Reading statements into workbooks --------------------------------------------------

def start_fresh() -> None:
    """Forget every workbook and statement of this visit."""
    shutil.rmtree(settings.output_dir, ignore_errors=True)
    shutil.rmtree(settings.uploads_dir, ignore_errors=True)
    st.session_state.pop("last_workbook", None)
    st.session_state.pop("opened", None)


def process_upload(name: str, data: bytes, password: str, add_to: list[Path]) -> None:
    settings.uploads_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = settings.uploads_dir / f"{sha256_of_bytes(data)[:12]}_{name}"
    pdf_path.write_bytes(data)

    run_settings = dataclasses.replace(settings, pdf_passwords=[p for p in [password, *settings.pdf_passwords] if p])
    status = st.empty()
    if pdf_path.suffix.lower() in SPREADSHEET_TYPES:
        results = parse_spreadsheet(pdf_path, layouts=layouts, generic=generic, client_rules=client_rules,
                                    settings=run_settings)
    else:
        results = parse_document(
            pdf_path,
            layouts=layouts,
            generic=generic,
            client_rules=client_rules,
            settings=run_settings,
            interactive=False,
            progress=lambda message: status.info(f"**{name}**: {message}"),
        )
    status.empty()
    if len(results) > 1:
        st.write(f"**{name}** holds {len(results)} statements - each is checked on its own:")
    for result in results:
        add_statement(result, add_to)


def _workbook_for(result, add_to: list[Path]) -> Path:
    """Each statement gets a workbook of its own - statements are never
    combined - unless a workbook of its account was opened under "Continue a
    workbook" to add it to."""
    for book in add_to:
        rows = read_transactions(book) if book.exists() else []
        if rows and str(rows[0].get("Account") or "") == (result.account_number or "") and result.account_number:
            return book
    return statement_workbook_path(settings.output_dir, result.account_number, result.statement_period,
                                   result.source_file)


def add_statement(result, add_to: list[Path]) -> bool:
    """Writes one statement to its workbook and shows its status. Every
    readable statement is added; anything that couldn't be confirmed is
    marked REVIEW_REQUIRED rather than turned away. True if added."""
    name = result.source_file
    if not result.ok:
        st.error(f"**{name}** - no transactions could be read, so nothing was added. {result.error}")
        return False

    workbook = _workbook_for(result, add_to)
    try:
        written = append_transactions(workbook, result.transactions, categories,
                                      statement=statement_review_row(result))
    except (WorkbookLockedError, MixedAccountsError) as exc:
        st.error(str(exc))
        return False
    st.session_state["last_workbook"] = str(workbook)

    report = result.report or {}
    dupes = f" ({written.skipped_duplicates} were already in the workbook)" if written.skipped_duplicates else ""
    review = [t for t in result.transactions if t.status != "APPROVED"]
    headline = (f"**{name}** - {result.bank_display_name}, {result.statement_period}: "
                f"{written.added} transactions added to **{workbook.stem}**{dupes}.")
    figures = ""
    if report.get("opening_balance") is not None:
        figures = (f" Opening {report['opening_balance']:,.2f} + credits {report['total_credits']:,.2f} "
                   f"- debits {report['total_debits']:,.2f} = {report['computed_closing']:,.2f}")
        if report.get("printed_closing") is not None:
            figures += f" (statement's closing balance: {report['printed_closing']:,.2f})"
        figures += "."
    if result.status == "APPROVED":
        st.success(f"APPROVED - {headline} Every row is confirmed by the running balance and by a second, "
                   f"independent reading.{figures}")
    else:
        issues = "".join(f"\n- {p}" for p in result.problems)
        st.warning(f"REVIEW_REQUIRED - {headline} {len(review)} row(s) need checking against the statement "
                   f"(highlighted in the workbook, and listed under **Checks**).{figures}{issues}")
    if result.warning:
        st.caption(f"{name}: {result.warning}")
    return True


# --- Look and helpers ---------------------------------------------------------------------

st.markdown("""
<style>
  .block-container {padding-top: 2.2rem; max-width: 1280px;}
  [data-testid="stMetricValue"] {font-size: 1.45rem; font-weight: 600;}
  [data-testid="stMetricLabel"] p {font-size: 0.85rem; color: #5B6575; text-transform: uppercase;
                                   letter-spacing: .03em;}
  [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h2 {margin-bottom: 0;}
  div[data-testid="stCaptionContainer"] p {color: #5B6575;}
</style>
""", unsafe_allow_html=True)


def rand(value) -> str:
    """R 47 436.83 - South African style, a space between thousands."""
    if value is None or (isinstance(value, float) and value != value):
        return "-"
    return ("-" if value < -0.004 else "") + "R " + f"{abs(value):,.2f}".replace(",", " ")


def page_header(title: str, subtitle: str = "") -> None:
    st.header(title, anchor=False)
    if subtitle:
        st.caption(subtitle)


def figures_row(items: list[tuple[str, str, str | None]]) -> None:
    """A row of figure cards: (label, value, help)."""
    for column, (label, value, help_text) in zip(st.columns(len(items)), items):
        with column.container(border=True):
            st.metric(label, value, help=help_text)


def money_table(df: pd.DataFrame, money: list[str], **kwargs) -> None:
    config = {c: st.column_config.NumberColumn(c, format="accounting") for c in money if c in df}
    st.dataframe(df, column_config=config, use_container_width=True, **kwargs)


# --- Pages ---------------------------------------------------------------------------------

def upload_page() -> None:
    page_header("Upload statements", "Bank statements in, workbooks and reports out - any South African bank.")
    opened = [Path(p) for p in st.session_state.get("opened", []) if Path(p).exists()]
    with st.form("upload", clear_on_submit=True):
        files = st.file_uploader("Bank statement PDFs - digital or scanned", type="pdf",
                                 accept_multiple_files=True)
        sheets = st.file_uploader(
            "Bank statement spreadsheets - Excel (.xlsx) or CSV",
            type=["xlsx", "xlsm", "csv"], accept_multiple_files=True,
            help="A statement exported or typed into a spreadsheet: a header row with a date column and an "
                 "amount (or debit and credit) column, and a balance column. Titles, notes, totals and extra "
                 "columns around it are fine. Older .xls files: save them as .xlsx first.")
        password = st.text_input(
            "PDF password (only if the statements are locked)",
            type="password",
            help="Used for this upload only. To stop typing it, add it to STATEMENT_PDF_PASSWORDS in .env.",
        )
        add_to_opened = st.checkbox(
            f"Add them to the workbook I opened ({', '.join(p.stem for p in opened)})", value=True,
        ) if opened else False
        submitted = st.form_submit_button("Read statements", type="primary", icon=":material/play_arrow:")
    st.caption("Each statement gets a workbook of its own - statements are never combined, and nothing from "
               "earlier uploads is carried over. Every figure is kept exactly as on the statement and checked "
               "against its running balance; anything that can't be confirmed is marked for review, never "
               "guessed. To add statements to a workbook you downloaded before, open it under "
               "**Continue a workbook**.")
    if not submitted:
        return
    uploads = list(files or []) + list(sheets or [])
    if not uploads:
        st.warning("Choose at least one PDF or spreadsheet first.")
        return
    if not add_to_opened:
        start_fresh()
    for f in uploads:
        with st.spinner(f"Reading {f.name}..."):
            process_upload(f.name, f.getvalue(), password, opened if add_to_opened else [])
    if list_workbooks(settings.output_dir):
        st.info("Done. See **Overview**, **Monthly**, **VAT** and **Checks** in the menu, and **Download** the "
                "workbooks.", icon=":material/task_alt:")


def continue_page() -> None:
    page_header("Continue a workbook", "Open a workbook you downloaded before, to add statements to it.")
    with st.form("continue", clear_on_submit=True):
        uploaded = st.file_uploader("Workbook you downloaded from this page before", type="xlsx",
                                    accept_multiple_files=True,
                                    help="Only a workbook this page made. Bank statements go under Upload statements.")
        submitted = st.form_submit_button("Open workbook", type="primary", icon=":material/folder_open:")
    st.caption("The statements you then upload under **Upload statements** are added to it - a statement already "
               "in it is never added twice.")
    if not submitted:
        return
    if not uploaded:
        st.warning("Choose a workbook first.")
        return
    start_fresh()
    opened = []
    for w in uploaded:
        target = restore_workbook(w.getvalue(), settings.output_dir, w.name)
        if target is None:
            st.error(f"**{w.name}** isn't a workbook downloaded from this page, so it was not opened.")
            continue
        opened.append(str(target))
        st.session_state["last_workbook"] = str(target)
        st.success(f"Opened **{target.stem}** ({len(read_transactions(target))} transactions). "
                   "Go to **Upload statements** to add statements to it.")
    st.session_state["opened"] = opened


def _frame(rows: list[dict]) -> pd.DataFrame:
    df = pd.DataFrame(rows)
    for col in ("Debit", "Credit"):
        df[col] = pd.to_numeric(df[col]).fillna(0.0)
    df["Balance"] = pd.to_numeric(df["Balance"])
    # Same category types the workbook's reports use (incl. hand-typed categories).
    df["Type"] = df["Category"].map({c.name: c.type for c in report_categories(categories, rows)})
    return df


def _opening_closing(rows: list[dict]) -> tuple[float | None, float | None]:
    """The balance before the first transaction and after the last, from the
    statements' own running balance."""
    with_balance = [r for r in rows if r.get("Balance") is not None]
    if not with_balance:
        return None, None
    first = with_balance[0]
    opening = first["Balance"] + (first.get("Debit") or 0) - (first.get("Credit") or 0)
    # Rows of the first transaction without their own balance come before it.
    for r in rows[:rows.index(first)]:
        opening += (r.get("Debit") or 0) - (r.get("Credit") or 0)
    return opening, with_balance[-1]["Balance"]


def monthly_table(rows: list[dict]) -> pd.DataFrame:
    """One line per month: balances, money in and out, income, expenses,
    profit and VAT - the same figures as the workbook's sheets."""
    types = {c.name: c.type for c in report_categories(categories, rows)}
    vat = {m["Month"]: m for m in vat_by_month(rows, categories)}
    by_month: dict[str, list[dict]] = defaultdict(list)
    for r in rows:
        if r.get("Month"):
            by_month[r["Month"]].append(r)
    out = []
    for month in sorted(by_month):
        part = by_month[month]
        opening, closing = _opening_closing(part)
        income = sum((r.get("Credit") or 0) - (r.get("Debit") or 0) for r in part if types.get(r["Category"]) == "income")
        expenses = sum((r.get("Debit") or 0) - (r.get("Credit") or 0) for r in part
                       if types.get(r["Category"]) == "expense")
        out.append({
            "Month": month_label(month),
            "Opening balance": opening,
            "Money in": sum(r.get("Credit") or 0 for r in part),
            "Money out": sum(r.get("Debit") or 0 for r in part),
            "Closing balance": closing,
            "Income": income,
            "Expenses": expenses,
            "Net profit": income - expenses,
            "VAT payable": vat[month]["VAT payable"] if month in vat else 0.0,
            "Transactions": len(part),
            "To review": sum(1 for r in part if r.get("Status") == "REVIEW_REQUIRED"),
        })
    return pd.DataFrame(out)


MONEY_COLUMNS = ["Opening balance", "Money in", "Money out", "Closing balance", "Income", "Expenses", "Net profit",
                 "VAT payable"]


def overview_page(rows: list[dict], period: Period) -> None:
    page_header("Overview", f"{period.label} - from {len(rows)} transactions, exactly as on the statements.")
    df = _frame(rows)
    income = df[df["Type"] == "income"]
    expenses = df[df["Type"] == "expense"]
    total_income = income["Credit"].sum() - income["Debit"].sum()
    total_expenses = expenses["Debit"].sum() - expenses["Credit"].sum()
    opening, closing = _opening_closing(rows)
    vat = sum(m["VAT payable"] for m in vat_by_month(rows, categories))
    figures_row([("Opening balance", rand(opening), "Before the period's first transaction"),
                 ("Money in", rand(df["Credit"].sum()), None),
                 ("Money out", rand(df["Debit"].sum()), None),
                 ("Closing balance", rand(closing), "After the period's last transaction")])
    figures_row([("Income", rand(total_income), None),
                 ("Expenses", rand(total_expenses), None),
                 ("Net profit", rand(total_income - total_expenses), "Income minus expenses; transfers and "
                                                                        "cash withdrawals aren't expenses"),
                 ("VAT payable", rand(vat), f"Calculated at {VAT_RATE:.0%} - not read from the statements")])

    left, right = st.columns(2)
    with left, st.container(border=True):
        st.markdown("**Money in and out by month**")
        monthly = df.groupby("Month")[["Credit", "Debit"]].sum().rename(
            columns={"Credit": "Money in", "Debit": "Money out"})
        monthly.index = [month_label(m) for m in monthly.index]
        st.bar_chart(monthly, stack=False, color=["#2E7D5B", "#B23A3A"])
    with right, st.container(border=True):
        st.markdown("**Spending by category**")
        spend = defaultdict(float)
        for _, r in expenses.iterrows():
            spend[r["Category"]] += r["Debit"] - r["Credit"]
        if spend:
            st.bar_chart(pd.Series(spend, name="Rand").sort_values(ascending=False), horizontal=True,
                         color="#1F3A68")
        else:
            st.caption("No expenses in this period.")

    # Income statement: each category's net per month (income as money in,
    # everything else as money out); transfers and drawings apart, outside profit.
    df["Net"] = df["Credit"] - df["Debit"]
    df.loc[df["Type"] != "income", "Net"] *= -1
    for title, kinds in (("Income statement", ("income", "expense")),
                         ("Not in profit: transfers, cash and drawings (money out)", ("transfer", "drawings"))):
        part = df[df["Type"].isin(kinds)]
        if part.empty:
            continue
        st.subheader(title, anchor=False)
        table = part.pivot_table(index=["Type", "Category"], columns="Month", values="Net", aggfunc="sum",
                                 fill_value=0.0)
        table.columns = [month_label(m) for m in table.columns]
        table["Total"] = table.sum(axis=1)
        money_table(table.round(2), list(table.columns))
    st.caption("Categories come from config/categories.yaml, your own edits in the workbook, and suggestions "
               "learnt from them - see Checks.")


def monthly_page(rows: list[dict], period: Period) -> None:
    page_header("Monthly", f"Each month of {period.label} - the workbook's Monthly Summary and Period Summary "
                           "sheets have the same figures.")
    table = monthly_table(rows)
    if table.empty:
        st.info("No dated transactions in this period.")
        return
    totals = {c: table[c].sum() for c in ("Money in", "Money out", "Income", "Expenses", "Net profit",
                                          "VAT payable", "Transactions", "To review")}
    totals.update({"Month": "Total", "Opening balance": table["Opening balance"].iloc[0],
                   "Closing balance": table["Closing balance"].iloc[-1]})
    shown = pd.concat([table, pd.DataFrame([totals])], ignore_index=True)
    money_table(shown.round(2), MONEY_COLUMNS, hide_index=True)

    st.subheader("One month in detail", anchor=False)
    month = st.selectbox("Month", list(table["Month"]), index=len(table) - 1)
    key = next(m for m in period.months if month_label(m) == month)
    part = [r for r in rows if r.get("Month") == key]
    line = table[table["Month"] == month].iloc[0]
    figures_row([("Money in", rand(line["Money in"]), None), ("Money out", rand(line["Money out"]), None),
                 ("Net profit", rand(line["Net profit"]), None), ("Closing balance", rand(line["Closing balance"]),
                                                                  None)])
    df = _frame(part)
    by_category = df.groupby(["Type", "Category"])[["Credit", "Debit"]].sum().rename(
        columns={"Credit": "Money in", "Debit": "Money out"}).reset_index()
    money_table(by_category.round(2), ["Money in", "Money out"], hide_index=True)
    with st.expander(f"All {len(part)} transactions of {month}"):
        money_table(df[[c for c in ("Date", "Description", "Category", "Debit", "Credit", "Balance", "Status")
                        if c in df]], ["Debit", "Credit", "Balance"], hide_index=True)


def vat_page(rows: list[dict], all_rows: list[dict], period: Period) -> None:
    page_header("VAT", f"{period.label}. VAT is calculated at {VAT_RATE:.0%} on transactions marked VAT Yes - "
                       "never read from the statements.")
    months = vat_by_month(rows, categories)
    if not months:
        st.info("No income or expense transactions in this period.")
        return
    table = pd.DataFrame(months)
    table["Month"] = [month_label(m) for m in table["Month"]]
    table = table.set_index("Month")[["Income", "VAT on income", "Expenses", "VAT on expenses", "VAT payable"]]
    totals = table.sum()  # added up unrounded, as the workbook does
    figures_row([("VAT on income (output)", rand(totals["VAT on income"]), None),
                 ("VAT on expenses (input)", rand(totals["VAT on expenses"]), None),
                 ("VAT payable", rand(totals["VAT payable"]), "Negative means a refund")])
    table.loc["Total"] = totals
    money_table(table.round(2), list(table.columns))

    st.subheader("VAT201 return", anchor=False)
    st.caption("The SARS VAT201 fields that bank statements can support. Claim input tax only where you hold a "
               "valid tax invoice; capital goods belong in field 14, and income without VAT goes in field 2 "
               "(zero-rated) or 3 (exempt).")
    all_months = [r.get("Month") for r in all_rows]
    left, right = st.columns([1, 2])
    category = left.selectbox("Your VAT category", ["A", "B", "C"],
                              format_func=lambda c: {"A": "A - two-monthly (Jan, Mar, May...)",
                                                     "B": "B - two-monthly (Feb, Apr, Jun...)",
                                                     "C": "C - monthly"}[c])
    choices = vat_periods(all_months, category)
    chosen = right.selectbox("VAT period", choices, index=len(choices) - 1, format_func=lambda p: p.label + (
        f" (only {len(p.months)} of {p.length} months uploaded)" if p.partial else ""))
    in_period = [r for r in all_rows if r.get("Month") in chosen.months]
    fields = vat201_by_month(vat_by_month(in_period, categories))
    labels = {field: label for field, label, _ in VAT201_FIELDS}
    money_table(pd.DataFrame([{"Field": field, "Description": labels[field],
                               "Amount": round(sum(m[field] for m in fields), 2)} for field in labels]),
                ["Amount"], hide_index=True)

    with st.expander("Transactions and their VAT setting"):
        df = _frame(rows)
        money_table(df[df["Type"].isin(["income", "expense"])][
            [c for c in ("Date", "Description", "Category", "VAT", "Debit", "Credit") if c in df]],
            ["Debit", "Credit"], hide_index=True)
    st.caption("Change a transaction's VAT Yes/No in the workbook's Transactions sheet. The workbook has the same "
               "figures per month, in VAT Summary and VAT201.")


def checks_page(rows: list[dict], selected: Path) -> None:
    page_header("Checks", "How each statement reconciled, and anything that needs a person to look at it - for "
                          "the whole workbook.")
    df = _frame(rows)
    reviews = read_reviews(selected)
    to_review = df[df["Status"] == "REVIEW_REQUIRED"] if "Status" in df else df.iloc[0:0]
    gaps = workbook_gaps(rows)
    figures_row([("Transactions", f"{len(df):,}".replace(",", " "), None),
                 ("Confirmed", f"{len(df) - len(to_review):,}".replace(",", " "),
                  "Confirmed by the running balance and by a second, independent reading"),
                 ("To review", f"{len(to_review):,}".replace(",", " "), None),
                 ("Balance gaps", str(len(gaps)), "Places where the balance doesn't carry on from one "
                                                  "transaction to the next")])
    if reviews:
        st.subheader("Statements", anchor=False)
        st.dataframe(pd.DataFrame(reviews)[[c for c in ("Source File", "Statement Period", "Status",
                                                         "Rows Needing Review", "Issues") if c in reviews[0]]],
                     hide_index=True, use_container_width=True)

    # Whole-workbook check: the running balance must carry on unbroken from one
    # transaction to the next, across every statement added.
    if gaps:
        st.error(f"**The balances don't join up in {len(gaps)} place(s)** - usually a missing statement for "
                 "that period, or a transaction edited in Excel. Totals for those months are incomplete.",
                 icon=":material/link_off:")
        money_table(pd.DataFrame({
            "After": g.after_date, "At": g.at_date, "Transaction": g.description,
            "Expected balance": g.expected, "Statement balance": g.actual,
            "Missing amount": round(g.actual - g.expected, 2),
        } for g in gaps), ["Expected balance", "Statement balance", "Missing amount"], hide_index=True)
    else:
        st.success("Balance check: every transaction follows on from the one before - no gaps or misreads.",
                   icon=":material/check_circle:")

    if len(to_review):
        st.warning(f"**{len(to_review)} row(s) need checking** against the statements - highlighted on the "
                   "Transactions sheet. Nothing was guessed: each is kept exactly as printed, with the reason. "
                   "Amounts whose direction isn't shown are in no total until confirmed.", icon=":material/flag:")
        money_table(to_review[[c for c in ("Date", "Description", "Debit", "Credit", "In/Out Not Shown", "Balance",
                                           "Review Reason", "Evidence", "Source File") if c in df]],
                    ["Debit", "Credit", "In/Out Not Shown", "Balance"], hide_index=True)
    else:
        st.success("Every row is confirmed by the running balance and by a second, independent reading.",
                   icon=":material/verified:")

    st.subheader("Categories", anchor=False)
    uncategorised = df[df["Category"].str.contains("uncategorised", na=False)]
    if len(uncategorised):
        st.warning(f"{len(uncategorised)} transaction(s) are uncategorised. Type a category into the Category "
                   "column in Excel, or add a rule in config/categories.yaml and click **Re-apply category rules**.")
        money_table(uncategorised[["Date", "Description", "Debit", "Credit"]], ["Debit", "Credit"], hide_index=True)
    suggested = (df[df["Category Source"].astype(str).str.startswith("Suggested")] if "Category Source" in df
                 else df.iloc[0:0])
    if len(suggested):
        st.info(f"{len(suggested)} transaction(s) have a **suggested** category, learnt from similar transactions "
                "whose category you typed. Check them in the Category column in Excel: type over any that are "
                "wrong, and later suggestions learn from it.")
        money_table(suggested[["Date", "Description", "Category", "Category Source", "Debit", "Credit"]],
                    ["Debit", "Credit"], hide_index=True)
    if not len(uncategorised) and not len(suggested):
        st.success("Every transaction has a category from a rule or from you.")
    if st.button("Re-apply category rules", icon=":material/refresh:"):
        try:
            append_transactions(selected, [], categories)
        except WorkbookLockedError as exc:
            st.error(str(exc))
        else:
            st.rerun()


def download_page(rows: list[dict], selected: Path) -> None:
    page_header("Download", "Excel workbooks with every report. Each cell holds its number (for phones and "
                            "previews) and its formula (so Excel recalculates your edits).")
    books = list_workbooks(settings.output_dir)
    if len(books) > 1:
        bundle = io.BytesIO()
        with zipfile.ZipFile(bundle, "w", zipfile.ZIP_DEFLATED) as z:
            for book in books:
                z.write(book, book.name)
        with st.container(border=True):
            st.markdown(f"**All {len(books)} workbooks** - one per statement")
            st.download_button(f"Download all {len(books)} workbooks (.zip)", data=bundle.getvalue(),
                               file_name="statement workbooks.zip", mime="application/zip", type="primary",
                               icon=":material/folder_zip:")
    periods_in = sorted({str(r.get("Statement Period")) for r in rows if r.get("Statement Period")})
    with st.container(border=True):
        st.markdown(f"**{selected.stem}**  \n{len(rows)} transactions from {len(periods_in)} statement "
                    f"period(s): {', '.join(periods_in)}")
        st.download_button(
            "Download this workbook",
            data=selected.read_bytes(),
            file_name=selected.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary" if len(books) <= 1 else "secondary",
            icon=":material/download:",
        )
    st.caption("Sheets: Review, Monthly Summary, Period Summary (months, quarters and financial years), Income "
               "Statement, Cash Flow, Category Breakdown, VAT Summary, VAT201, one VAT sheet per month, "
               "Transactions and Categories. Keep it: to add next month's statements, open it under "
               "**Continue a workbook**.")


# --- Menu ------------------------------------------------------------------------------------

UPLOAD, CONTINUE, OVERVIEW, MONTHLY, VAT, CHECKS, DOWNLOAD = (
    "Upload statements", "Continue a workbook", "Overview", "Monthly", "VAT", "Checks", "Download")
ICONS = {UPLOAD: ":material/upload_file:", CONTINUE: ":material/folder_open:", OVERVIEW: ":material/monitoring:",
         MONTHLY: ":material/calendar_month:", VAT: ":material/receipt_long:", CHECKS: ":material/fact_check:",
         DOWNLOAD: ":material/download:"}
CUSTOM = "Custom range of months"

with st.sidebar:
    st.markdown("## Bank Statement Tool")
    st.caption("Statements to financial reports")
    page = st.radio("Menu", list(ICONS), key="page", format_func=lambda p: f"{ICONS[p]}  {p}",
                    label_visibility="collapsed")
    st.divider()
account_slot = st.sidebar.container()  # filled once this run's uploads are in


def choose_account() -> Path | None:
    books = list_workbooks(settings.output_dir)
    with account_slot:
        if not books:
            st.caption("No workbook yet - upload statements to start one.")
            return None
        last = st.session_state.get("last_workbook")
        return st.selectbox("Workbook", books, index=next((i for i, b in enumerate(books) if str(b) == last), 0),
                            format_func=lambda p: p.stem)


def choose_period(rows: list[dict]) -> Period:
    """The reporting period picked in the side menu: all months, a financial
    year, a quarter, a month or a custom range."""
    months = sorted({r["Month"] for r in rows if r.get("Month")})
    options = periods(months)
    if not options:
        return Period("All transactions", "all", (), 0)
    labels = [p.label for p in options] + [CUSTOM]
    remembered = st.session_state.get("period_label")
    with account_slot:
        label = st.selectbox("Reporting period", labels,
                             index=labels.index(remembered) if remembered in labels else 0)
        st.session_state["period_label"] = label
        if label != CUSTOM:
            return next(p for p in options if p.label == label)
        start = st.selectbox("From", months, format_func=month_label)
        end = st.selectbox("To", months, index=len(months) - 1, format_func=month_label)
        if end < start:
            start, end = end, start
        chosen = tuple(m for m in months if start <= m <= end)
        return Period(f"{month_label(start)} - {month_label(end)}", "custom", chosen, len(chosen))


if page == UPLOAD:
    upload_page()
    choose_account()
elif page == CONTINUE:
    continue_page()
    choose_account()
else:
    selected = choose_account()
    rows = []
    if selected is not None:
        try:
            rows = read_transactions(selected)
        except (WorkbookLockedError, PermissionError):
            st.warning("The workbook is open in Excel. Close it there and refresh this page.")
            st.stop()
    if not rows:
        page_header(page)
        st.info("No transactions yet - upload statements under **Upload statements** in the menu.",
                icon=":material/upload_file:")
        st.stop()
    if page in (OVERVIEW, MONTHLY, VAT):
        period = choose_period(rows)
        in_period = rows if period.kind == "all" else [r for r in rows if r.get("Month") in period.months]
        if not in_period:
            page_header(page)
            st.info("No transactions in this period.")
            st.stop()
        if page == OVERVIEW:
            overview_page(in_period, period)
        elif page == MONTHLY:
            monthly_page(in_period, period)
        else:
            vat_page(in_period, rows, period)
    elif page == CHECKS:
        checks_page(rows, selected)
    else:
        download_page(rows, selected)
