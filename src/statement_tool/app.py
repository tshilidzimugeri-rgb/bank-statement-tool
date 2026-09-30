"""Upload page: drop in statement PDFs, get the workbook and reports.

Runs on this computer (run_app.bat) or hosted online (Streamlit Community
Cloud). Online it needs a password (APP_PASSWORD in the app's secrets).

A side menu splits it into pages: Upload statements, Continue a workbook,
Financials, VAT, Checks and Download. Each visit works in its own temporary
folder, and each upload makes a new workbook from just the statements
uploaded then - statements from earlier uploads never carry over. To add
statements to a workbook downloaded before, open it on "Continue a
workbook" first.

Run with run_app.bat, or:  .venv\\Scripts\\streamlit run src\\statement_tool\\app.py
"""
from __future__ import annotations

import dataclasses
import hmac
import os
import shutil
import sys
import tempfile
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
    vat_by_month,
    workbook_path_for,
)
from statement_tool.extract.document import parse_document
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


def process_upload(name: str, data: bytes, password: str) -> None:
    settings.uploads_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = settings.uploads_dir / f"{sha256_of_bytes(data)[:12]}_{name}"
    pdf_path.write_bytes(data)

    run_settings = dataclasses.replace(settings, pdf_passwords=[p for p in [password, *settings.pdf_passwords] if p])
    status = st.empty()
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
        add_statement(result)


def add_statement(result) -> bool:
    """Writes one statement to its account's workbook and shows its status.
    Every readable statement is added; anything that couldn't be confirmed is
    marked REVIEW_REQUIRED rather than turned away. True if added."""
    name = result.source_file
    if not result.ok:
        st.error(f"**{name}** - no transactions could be read, so nothing was added. {result.error}")
        return False

    workbook = workbook_path_for(settings.output_dir, result.client, result.account_number, result.source_file)
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


# --- Pages ---------------------------------------------------------------------------------

def upload_page() -> None:
    st.header("Upload statements")
    opened = [Path(p) for p in st.session_state.get("opened", []) if Path(p).exists()]
    st.info("Each upload makes a **new workbook** from just the statements you upload now - statements "
            "from earlier uploads are never carried over. To add statements to a workbook you downloaded "
            "before, open it first under **Continue a workbook** in the menu.")
    with st.form("upload", clear_on_submit=True):
        files = st.file_uploader("Bank statement PDFs (any bank, digital or scanned)", type="pdf",
                                 accept_multiple_files=True)
        password = st.text_input(
            "PDF password (only if the statements are locked)",
            type="password",
            help="Used for this upload only. To stop typing it, add it to STATEMENT_PDF_PASSWORDS in .env.",
        )
        add_to_opened = st.checkbox(
            f"Add them to the workbook I opened ({', '.join(p.stem for p in opened)})", value=True,
        ) if opened else False
        submitted = st.form_submit_button("Read statements", type="primary")
    if not submitted:
        return
    if not files:
        st.warning("Choose at least one PDF first.")
        return
    if not add_to_opened:
        start_fresh()
    for f in files:
        with st.spinner(f"Reading {f.name}..."):
            process_upload(f.name, f.getvalue(), password)
    if list_workbooks(settings.output_dir):
        st.info("Done. Use the menu for **Financials**, **VAT**, **Checks** and **Download**.")


def continue_page() -> None:
    st.header("Continue a workbook")
    st.write("Open a workbook you downloaded from this page before. The statements you then upload under "
             "**Upload statements** are added to it (a statement already in it is not added twice), and "
             "Financials, VAT, Checks and Download show it.")
    with st.form("continue", clear_on_submit=True):
        uploaded = st.file_uploader("Workbook you downloaded from this page before", type="xlsx",
                                    accept_multiple_files=True,
                                    help="Only a workbook this page made. Bank statements go under Upload statements.")
        submitted = st.form_submit_button("Open workbook", type="primary")
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
    df["Debit"] = pd.to_numeric(df["Debit"]).fillna(0.0)
    df["Credit"] = pd.to_numeric(df["Credit"]).fillna(0.0)
    # Same category types the workbook's reports use (incl. hand-typed categories).
    df["Type"] = df["Category"].map({c.name: c.type for c in report_categories(categories, rows)})
    return df


def financials_page(rows: list[dict]) -> None:
    df = _frame(rows)
    income = df[df["Type"] == "income"]
    expenses = df[df["Type"] == "expense"]
    total_income = income["Credit"].sum() - income["Debit"].sum()
    total_expenses = expenses["Debit"].sum() - expenses["Credit"].sum()
    st.header(f"Financials {df['Month'].min()} to {df['Month'].max()}")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Income", f"R {total_income:,.2f}")
    c2.metric("Expenses", f"R {total_expenses:,.2f}")
    c3.metric("Net profit", f"R {total_income - total_expenses:,.2f}")
    balances = df.dropna(subset=["Balance"])
    c4.metric("Latest balance", f"R {balances['Balance'].iloc[-1]:,.2f}" if len(balances) else "-")

    left, right = st.columns(2)
    with left:
        st.markdown("**Money in vs out by month**")
        monthly = df.groupby("Month")[["Credit", "Debit"]].sum().rename(
            columns={"Credit": "Money in", "Debit": "Money out"})
        st.bar_chart(monthly, stack=False)
    with right:
        st.markdown("**Spending by category**")
        spend = defaultdict(float)
        for _, r in expenses.iterrows():
            spend[r["Category"]] += r["Debit"] - r["Credit"]
        st.bar_chart(pd.Series(spend, name="Rand").sort_values(ascending=False), horizontal=True)

    # Income statement: each category's net per month (income as money in,
    # everything else as money out); transfers and drawings apart, outside profit.
    df["Net"] = df["Credit"] - df["Debit"]
    df.loc[df["Type"] != "income", "Net"] *= -1
    for title, kinds in (("Income statement", ("income", "expense")),
                         ("Not profit: transfers and drawings (money out)", ("transfer", "drawings"))):
        part = df[df["Type"].isin(kinds)]
        if part.empty:
            continue
        st.subheader(title)
        table = part.pivot_table(index=["Type", "Category"], columns="Month", values="Net", aggfunc="sum",
                                 fill_value=0.0)
        table["Total"] = table.sum(axis=1)
        st.dataframe(table.round(2), use_container_width=True)
    st.caption("Categories come from config/categories.yaml, your own edits in the workbook, and suggestions "
               "learnt from them - see Checks. Amounts are exactly as printed on the statements.")


def vat_page(rows: list[dict]) -> None:
    st.header("VAT")
    st.caption(f"VAT is **calculated** at {VAT_RATE:.0%} on the transactions whose VAT is Yes - it is not read "
               "from the statements. VAT payable = VAT on income minus VAT on expenses; negative means a "
               "refund. Set each transaction's VAT Yes/No on the workbook's Transactions sheet.")
    months = vat_by_month(rows, categories)
    if not months:
        st.info("No income or expense transactions yet.")
        return
    table = pd.DataFrame(months).set_index("Month")
    totals = table.sum().round(2)
    c1, c2, c3 = st.columns(3)
    c1.metric("VAT on income", f"R {totals['VAT on income']:,.2f}")
    c2.metric("VAT on expenses", f"R {totals['VAT on expenses']:,.2f}")
    c3.metric("VAT payable", f"R {totals['VAT payable']:,.2f}")
    table.loc["TOTAL"] = totals
    st.dataframe(table, use_container_width=True)
    with st.expander("Transactions and their VAT setting"):
        df = _frame(rows)
        st.dataframe(df[df["Type"].isin(["income", "expense"])][
            [c for c in ("Date", "Description", "Category", "VAT", "Debit", "Credit") if c in df]], hide_index=True)
    st.caption("The workbook has the same figures per month (one sheet per month, and VAT Summary).")


def checks_page(rows: list[dict], selected: Path) -> None:
    st.header("Checks")
    df = _frame(rows)
    reviews = read_reviews(selected)
    if reviews:
        st.subheader("Statements")
        st.dataframe(pd.DataFrame(reviews)[[c for c in ("Source File", "Statement Period", "Status",
                                                         "Rows Needing Review", "Issues") if c in reviews[0]]],
                     hide_index=True)

    # Whole-workbook check: the running balance must carry on unbroken from one
    # transaction to the next, across every statement added.
    gaps = workbook_gaps(rows)
    if gaps:
        st.error(f"**The workbook's balances don't join up in {len(gaps)} place(s)** - usually a missing "
                 "statement for that period, or a transaction edited in Excel. Totals for those months are "
                 "incomplete.")
        st.dataframe(pd.DataFrame({
            "Account": g.account, "After": g.after_date, "At": g.at_date, "Transaction": g.description,
            "Expected balance": g.expected, "Statement balance": g.actual,
            "Missing amount": round(g.actual - g.expected, 2),
        } for g in gaps), hide_index=True)
    else:
        st.success("Balance check: every transaction follows on from the one before - no gaps or misreads.")

    to_review = df[df["Status"] == "REVIEW_REQUIRED"] if "Status" in df else df.iloc[0:0]
    if len(to_review):
        st.warning(f"**{len(to_review)} row(s) need checking** against the statements - they are highlighted on "
                   "the Transactions sheet. Nothing was guessed: each is kept exactly as printed, with the "
                   "reason. Amounts whose direction isn't shown are in no total until confirmed.")
        st.dataframe(to_review[[c for c in ("Date", "Description", "Debit", "Credit", "In/Out Not Shown",
                                            "Balance", "Review Reason", "Evidence", "Source File") if c in df]],
                     hide_index=True)
    else:
        st.success("Every row is APPROVED: confirmed by the running balance and by a second, independent reading.")

    st.subheader("Categories")
    uncategorised = df[df["Category"].str.contains("uncategorised", na=False)]
    if len(uncategorised):
        st.warning(f"{len(uncategorised)} transaction(s) are uncategorised. Add a rule for them in "
                   "config/categories.yaml and click **Re-apply category rules**, or type a category into the "
                   "Category column in Excel.")
        st.dataframe(uncategorised[["Date", "Description", "Debit", "Credit"]], hide_index=True)
    suggested = (df[df["Category Source"].astype(str).str.startswith("Suggested")] if "Category Source" in df
                 else df.iloc[0:0])
    if len(suggested):
        st.info(f"{len(suggested)} transaction(s) have a **suggested** category, learnt from similar transactions "
                "whose category you typed. Check them in the Category column in Excel: type over any that are "
                "wrong, and later suggestions learn from it.")
        st.dataframe(suggested[["Date", "Description", "Category", "Category Source", "Debit", "Credit"]],
                     hide_index=True)
    if not len(uncategorised) and not len(suggested):
        st.success("Every transaction has a category from a rule or from you.")
    if st.button("Re-apply category rules"):
        try:
            append_transactions(selected, [], categories)
        except WorkbookLockedError as exc:
            st.error(str(exc))
        else:
            st.rerun()


def download_page(rows: list[dict], selected: Path) -> None:
    st.header("Download")
    periods = sorted({str(r.get("Statement Period")) for r in rows if r.get("Statement Period")})
    st.write(f"**{selected.stem}**: {len(rows)} transactions from {len(periods)} statement period(s): "
             f"{', '.join(periods)}.")
    st.download_button(
        "Download Excel workbook",
        data=selected.read_bytes(),
        file_name=selected.name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
    st.caption("Sheets: Review, Monthly Summary, Income Statement, Cash Flow, Category Breakdown, VAT Summary, "
               "one VAT sheet per month, Transactions and Categories. Keep it: to add next month's statements, "
               "open it under **Continue a workbook**.")


# --- Menu ------------------------------------------------------------------------------------

UPLOAD, CONTINUE, FINANCIALS, VAT, CHECKS, DOWNLOAD = (
    "Upload statements", "Continue a workbook", "Financials", "VAT", "Checks", "Download")
st.sidebar.title("Bank Statement Tool")
page = st.sidebar.radio("Menu", [UPLOAD, CONTINUE, FINANCIALS, VAT, CHECKS, DOWNLOAD], key="page")
account_slot = st.sidebar.container()  # filled once this run's uploads are in


def choose_account() -> Path | None:
    books = list_workbooks(settings.output_dir)
    with account_slot:
        if not books:
            st.caption("No workbook yet - upload statements to start one.")
            return None
        last = st.session_state.get("last_workbook")
        return st.selectbox("Account", books, index=next((i for i, b in enumerate(books) if str(b) == last), 0),
                            format_func=lambda p: p.stem)


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
        st.info("No transactions yet - upload statements under **Upload statements** in the menu.")
        st.stop()
    if page == FINANCIALS:
        financials_page(rows)
    elif page == VAT:
        vat_page(rows)
    elif page == CHECKS:
        checks_page(rows, selected)
    else:
        download_page(rows, selected)
