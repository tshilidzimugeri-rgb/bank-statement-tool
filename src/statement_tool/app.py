"""Upload page: drop in statement PDFs, get the workbook and reports.

Runs on this computer (run_app.bat) or hosted online (Streamlit Community
Cloud). Online it needs a password (APP_PASSWORD in the app's secrets), and
keeps nothing between visits: each visit works in its own temporary folder,
so you upload your latest workbook together with new statements and
download it again when done.

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
    MixedAccountsError,
    WorkbookLockedError,
    append_transactions,
    list_workbooks,
    read_transactions,
    report_categories,
    restore_workbook,
    statement_review_row,
    workbook_path_for,
)
from statement_tool.extract.document import parse_document
from statement_tool.store import ProcessedStore, sha256_of_bytes

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

if HOSTED:
    # Each visit gets its own temporary folder: nothing is shared between
    # visitors or kept on the server afterwards.
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


def restore_upload(name: str, data: bytes) -> None:
    target = restore_workbook(data, settings.output_dir, name)
    if target is None:
        st.error(
            f"**{name}** isn't a workbook downloaded from this page, so it was not used. "
            "The top box is only for continuing a workbook from last time; bank statements go in the "
            "**Bank statement PDFs** box."
        )
        return
    st.session_state["last_workbook"] = str(target)
    st.info(f"Using your workbook **{target.stem}** ({len(read_transactions(target))} transactions).")


def process_upload(name: str, data: bytes, password: str, reprocess: bool) -> None:
    content_hash = sha256_of_bytes(data)
    source_id = f"upload:{content_hash}"
    with ProcessedStore(settings.processed_db) as store:
        if not HOSTED and store.is_processed(source_id, content_hash) and not reprocess:
            st.info(f"**{name}** was already added earlier - skipped.")
            return

        settings.uploads_dir.mkdir(parents=True, exist_ok=True)
        pdf_path = settings.uploads_dir / f"{content_hash[:12]}_{name}"
        pdf_path.write_bytes(data)

        run_settings = dataclasses.replace(
            settings, pdf_passwords=[p for p in [password, *settings.pdf_passwords] if p]
        )
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
        added = [add_statement(r) for r in results]
        if all(added):
            store.mark_processed(source_id, content_hash, name, results[0].client, results[0].bank_display_name,
                                 sum(len(r.transactions) for r in results))


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
                   f"(highlighted in the workbook).{figures}{issues}")
        if review:
            st.dataframe(pd.DataFrame(
                {"Date": t.date, "Description": t.description, "Debit": t.debit, "Credit": t.credit,
                 "In/out not shown": t.unassigned, "Balance": t.balance, "Why": t.check,
                 "As printed": t.evidence}
                for t in review), hide_index=True)
    if result.warning:
        st.caption(f"{name}: {result.warning}")
    return True


st.title("Bank Statement Tool")
if HOSTED:
    st.info("Online: each upload makes a new workbook from just the files you upload that time. "
            "To add to a workbook you downloaded earlier, upload it in the top box together with the new "
            "statements, and download the updated workbook before closing the page.")
else:
    st.caption(f"Workbooks (one per bank account): `{settings.output_dir}`")

with st.form("upload", clear_on_submit=True):
    existing = (
        st.file_uploader(
            "Workbook you downloaded from this page last time (optional)",
            type="xlsx",
            accept_multiple_files=True,
            help="Only a workbook this page made. Bank statements go in the box below.",
        )
        if HOSTED else []
    )
    files = st.file_uploader("Bank statement PDFs (any bank, digital or scanned)", type="pdf",
                             accept_multiple_files=True)
    password = st.text_input(
        "PDF password (only if the statements are locked)",
        type="password",
        help="Used for this upload only. To stop typing it, add it to STATEMENT_PDF_PASSWORDS in .env.",
    )
    reprocess = False if HOSTED else st.checkbox(
        "Process again even if already added (duplicates are still skipped)"
    )
    submitted = st.form_submit_button("Add to workbook", type="primary")

if submitted:
    if HOSTED and (files or existing):
        # Online, each upload starts from nothing: the workbook holds only the
        # statements (and any workbook) uploaded this time, not ones added
        # earlier in the same visit.
        shutil.rmtree(settings.output_dir, ignore_errors=True)
        shutil.rmtree(settings.uploads_dir, ignore_errors=True)
        st.session_state.pop("last_workbook", None)
    for w in existing or []:
        restore_upload(w.name, w.getvalue())
    if not files and not existing:
        st.warning("Choose at least one PDF first.")
    for f in files or []:
        with st.spinner(f"Reading {f.name}..."):
            process_upload(f.name, f.getvalue(), password, reprocess)

books = list_workbooks(settings.output_dir)
if not books:
    st.info("No transactions yet - upload a statement above.")
    st.stop()
last = st.session_state.get("last_workbook")
selected = st.selectbox(
    "Account",
    books,
    index=next((i for i, b in enumerate(books) if str(b) == last), 0),
    format_func=lambda p: p.stem,
)

try:
    rows = read_transactions(selected)
    workbook_bytes = selected.read_bytes() if rows else b""
except (WorkbookLockedError, PermissionError):
    st.warning("The workbook is open in Excel. Close it there and refresh this page to see the overview.")
    st.stop()
if not rows:
    st.info("No transactions yet - upload a statement above.")
    st.stop()

st.download_button(
    "Download Excel workbook",
    data=workbook_bytes,
    file_name=selected.name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

# Whole-workbook check: the running balance must carry on unbroken from one
# transaction to the next, across every statement added.
gaps = workbook_gaps(rows)
if gaps:
    st.error(
        f"**The workbook's balances don't join up in {len(gaps)} place(s)** - usually a missing statement "
        "for that period, or a transaction edited in Excel. Totals for those months are incomplete."
    )
    st.dataframe(
        pd.DataFrame(
            {
                "Account": g.account,
                "After": g.after_date,
                "At": g.at_date,
                "Transaction": g.description,
                "Expected balance": g.expected,
                "Statement balance": g.actual,
                "Missing amount": round(g.actual - g.expected, 2),
            }
            for g in gaps
        ),
        hide_index=True,
    )
else:
    st.success("Balance check: every transaction follows on from the one before - no gaps or misreads.")

# Quick on-screen view. The workbook's report sheets are the full version,
# and recalculate if anything is edited there.
df = pd.DataFrame(rows)
df["Debit"] = pd.to_numeric(df["Debit"]).fillna(0.0)
df["Credit"] = pd.to_numeric(df["Credit"]).fillna(0.0)
# Same category types the workbook's reports use (incl. hand-typed categories).
df["Type"] = df["Category"].map({c.name: c.type for c in report_categories(categories, rows)})

income = df[df["Type"] == "income"]
expenses = df[df["Type"] == "expense"]
total_income = income["Credit"].sum() - income["Debit"].sum()
total_expenses = expenses["Debit"].sum() - expenses["Credit"].sum()
first, last = df["Month"].min(), df["Month"].max()

st.subheader(f"Overview {first} to {last}")
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
        columns={"Credit": "Money in", "Debit": "Money out"}
    )
    st.bar_chart(monthly, stack=False)
with right:
    st.markdown("**Spending by category**")
    spend = defaultdict(float)
    for _, r in expenses.iterrows():
        spend[r["Category"]] += r["Debit"] - r["Credit"]
    st.bar_chart(pd.Series(spend, name="Rand").sort_values(ascending=False), horizontal=True)

if "Status" in df:
    to_review = df[df["Status"] == "REVIEW_REQUIRED"]
    if len(to_review):
        st.warning(f"**{len(to_review)} row(s) in this account need checking** against the statements - they "
                   "are highlighted on the Transactions sheet, and the Review sheet has each statement's "
                   "reconciliation. Amounts whose direction isn't shown are in no total until confirmed.")
        st.dataframe(to_review[[c for c in ("Date", "Description", "Debit", "Credit", "In/Out Not Shown",
                                            "Balance", "Review Reason", "Evidence", "Source File") if c in df]],
                     hide_index=True)
    else:
        st.success("Every row in this account is APPROVED: confirmed by the running balance and by a second, "
                   "independent reading.")

uncategorised = df[df["Category"].str.contains("uncategorised", na=False)]
if len(uncategorised):
    st.warning(
        f"{len(uncategorised)} transaction(s) are uncategorised. Add a rule for them in "
        "config/categories.yaml and click **Re-apply category rules**, or type a category into the "
        "Category column in Excel."
    )
    st.dataframe(uncategorised[["Date", "Description", "Debit", "Credit"]], hide_index=True)
    if st.button("Re-apply category rules"):
        try:
            append_transactions(selected, [], categories)
        except WorkbookLockedError as exc:
            st.error(str(exc))
        else:
            st.rerun()

with st.expander(f"All {len(df)} transactions"):
    st.dataframe(df[[c for c in ("Date", "Description", "Category", "VAT", "Debit", "Credit", "Balance", "Status")
                     if c in df]], hide_index=True)
