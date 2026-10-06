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
import itertools
import os
import re
import shutil
import sys
import tempfile
import zipfile
from collections import defaultdict
from functools import partial
from html import escape
from pathlib import Path

import altair as alt
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
from statement_tool.categorize import for_vat_registration, load_categories
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
    all_categories = load_categories(settings.categories_config)
    categories = for_vat_registration(all_categories, st.session_state.get("vat_registered", True))
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


def process_upload(name: str, data: bytes, password: str, add_to: list[Path], progress=None) -> list[dict]:
    """Reads one uploaded file into workbooks; a summary per statement in it."""
    settings.uploads_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = settings.uploads_dir / f"{sha256_of_bytes(data)[:12]}_{name}"
    pdf_path.write_bytes(data)

    run_settings = dataclasses.replace(settings, pdf_passwords=[p for p in [password, *settings.pdf_passwords] if p])
    if pdf_path.suffix.lower() in SPREADSHEET_TYPES:
        results = parse_spreadsheet(pdf_path, layouts=layouts, generic=generic, client_rules=client_rules,
                                    settings=run_settings, label=pdf_path.name)
    else:
        results = parse_document(
            pdf_path,
            layouts=layouts,
            generic=generic,
            client_rules=client_rules,
            settings=run_settings,
            interactive=False,
            progress=progress or (lambda message: None),
        )
    return [add_statement(result, add_to, name) for result in results]


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


def add_statement(result, add_to: list[Path], upload_name: str) -> dict:
    """Writes one statement to its workbook. Every readable statement is
    added; anything that couldn't be confirmed is marked REVIEW_REQUIRED
    rather than turned away. Returns what to show about it."""
    label = re.sub(r"^[0-9a-f]{12}_", "", result.source_file)
    summary = {"name": label, "file": upload_name, "ok": False, "error": None}
    if not result.ok:
        summary["error"] = f"No transactions could be read, so nothing was added. {result.error or ''}".strip()
        return summary
    workbook = _workbook_for(result, add_to)
    try:
        written = append_transactions(workbook, result.transactions, categories,
                                      statement=statement_review_row(result))
    except (WorkbookLockedError, MixedAccountsError) as exc:
        summary["error"] = str(exc)
        return summary
    st.session_state["last_workbook"] = str(workbook)
    summary.update(
        ok=True, status=result.status, bank=result.bank_display_name or "Unknown bank",
        period=result.statement_period or "Unknown", account=result.account_number or "not found",
        workbook=workbook.stem, added=written.added, duplicates=written.skipped_duplicates,
        review=sum(1 for t in result.transactions if t.status != "APPROVED"), rows=len(result.transactions),
        problems=list(result.problems), warning=result.warning, report=dict(result.report or {}))
    return summary


# --- Design system ------------------------------------------------------------------------
# A premium, calm fintech look: Inter type, a white surface on a cool grey
# canvas, indigo for action, green/amber/red only for status, 14px radius,
# soft layered shadows and gentle motion. Colours meet WCAG AA contrast.

html = partial(st.markdown, unsafe_allow_html=True)

html("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
:root {
  --ink: #0B1220; --ink-2: #334155; --muted: #5B6779; --line: #E4E7EC; --line-2: #EEF0F4;
  --canvas: #F6F7FB; --surface: #FFFFFF; --brand: #4F46E5; --brand-2: #6366F1; --brand-ink: #3730A3;
  --brand-soft: #EEF2FF; --ok: #067647; --ok-soft: #ECFDF3; --warn: #B54708; --warn-soft: #FFFAEB;
  --bad: #B42318; --bad-soft: #FEF3F2; --radius: 14px;
  --shadow: 0 1px 2px rgba(16,24,40,.04), 0 1px 3px rgba(16,24,40,.06);
  --shadow-lg: 0 12px 24px -8px rgba(16,24,40,.12), 0 4px 8px -4px rgba(16,24,40,.06);
}
html, body, [class*="css"], .stApp, button, input, textarea, select { font-family: 'Inter', system-ui, sans-serif; }
.stApp { background: var(--canvas); color: var(--ink); }
header[data-testid="stHeader"] { background: transparent; }
.block-container { max-width: 1200px; padding: 1.4rem 2.2rem 4rem; animation: fadeUp .35s ease both; }
@keyframes fadeUp { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: none; } }
@keyframes shimmer { 0% { background-position: -400px 0; } 100% { background-position: 400px 0; } }

/* Type scale */
h1, h2, h3 { color: var(--ink); letter-spacing: -0.02em; }
[data-testid="stHeading"] h2 { font-size: 1.75rem; font-weight: 700; margin: .2rem 0 0; }
[data-testid="stHeading"] h3 { font-size: 1.12rem; font-weight: 650; margin: 1.6rem 0 .2rem; }
[data-testid="stCaptionContainer"] p, .muted { color: var(--muted) !important; }

/* Cards: Streamlit's bordered containers become cards */
[class*="st-key-card_"] {
  background: var(--surface); border: 1px solid var(--line) !important; border-radius: var(--radius) !important;
  box-shadow: var(--shadow); padding: 1rem 1.1rem; transition: box-shadow .2s ease, transform .2s ease;
}
[class*="st-key-card_"]:hover { box-shadow: var(--shadow-lg); }

/* Figures */
[data-testid="stMetric"] { padding: .15rem .1rem; }
[data-testid="stMetricLabel"] p { font-size: .76rem; font-weight: 600; color: var(--muted); text-transform: uppercase;
  letter-spacing: .06em; }
[data-testid="stMetricValue"] { font-size: 1.6rem; font-weight: 700; color: var(--ink); letter-spacing: -0.02em; }

/* Sidebar navigation: each option is a nav item; groups get a small heading */
[data-testid="stSidebar"] { background: var(--surface); border-right: 1px solid var(--line); transition: none !important; }
[data-testid="stSidebarUserContent"] { padding-top: .25rem; }
[data-testid="stSidebar"] .stRadio, [data-testid="stSidebar"] [role="radiogroup"] { width: 100%; gap: 2px; align-items: stretch; }
[data-testid="stSidebar"] [role="radiogroup"] > div { position: relative; width: 100%; }
[data-testid="stSidebar"] [data-testid="stRadioOption"] {
  display: flex; align-items: center; width: 100%; margin: 0; padding: .55rem .75rem; border-radius: 10px;
  border: 1px solid transparent; cursor: pointer; transition: background .15s ease, border-color .15s ease;
}
[data-testid="stSidebar"] [data-testid="stRadioOption"] > div > div:not([data-testid="stMarkdownContainer"]) {
  display: none;  /* the radio dot */
}
[data-testid="stSidebar"] [data-testid="stRadioOption"] p { font-size: .92rem; font-weight: 500; color: var(--ink-2);
  display: flex; align-items: center; gap: .6rem; }
[data-testid="stSidebar"] [data-testid="stRadioOption"] [role="img"] { font-size: 1.15rem !important; color: #8A94A6; }
[data-testid="stSidebar"] [data-testid="stRadioOption"]:hover { background: var(--canvas); }
[data-testid="stSidebar"] [data-testid="stRadioOption"][data-selected="true"] {
  background: var(--brand-soft); border-color: #E0E7FF; box-shadow: inset 3px 0 0 var(--brand);
}
[data-testid="stSidebar"] [data-testid="stRadioOption"][data-selected="true"] p { color: var(--brand-ink); font-weight: 600; }
[data-testid="stSidebar"] [data-testid="stRadioOption"][data-selected="true"] [role="img"] { color: var(--brand); }
/* Groups: Data (upload, continue) - Reports (overview, monthly, VAT) - Review & export */
[data-testid="stSidebar"] [role="radiogroup"] > div:nth-child(1),
[data-testid="stSidebar"] [role="radiogroup"] > div:nth-child(3),
[data-testid="stSidebar"] [role="radiogroup"] > div:nth-child(6) { margin-top: 1.75rem; }
[data-testid="stSidebar"] [role="radiogroup"] > div:nth-child(1)::before,
[data-testid="stSidebar"] [role="radiogroup"] > div:nth-child(3)::before,
[data-testid="stSidebar"] [role="radiogroup"] > div:nth-child(6)::before {
  position: absolute; top: -1.3rem; left: .8rem; font-size: .68rem; font-weight: 700; letter-spacing: .09em;
  color: #8A94A6;
}
[data-testid="stSidebar"] [role="radiogroup"] > div:nth-child(1)::before { content: "DATA"; }
[data-testid="stSidebar"] [role="radiogroup"] > div:nth-child(3)::before { content: "REPORTS"; }
[data-testid="stSidebar"] [role="radiogroup"] > div:nth-child(6)::before { content: "REVIEW & EXPORT"; }

/* Buttons */
.stButton > button, .stDownloadButton > button, .stFormSubmitButton > button {
  border-radius: 10px; font-weight: 600; padding: .55rem 1.1rem; transition: transform .12s ease, box-shadow .2s;
}
.stButton > button:hover, .stDownloadButton > button:hover, .stFormSubmitButton > button:hover {
  transform: translateY(-1px); box-shadow: 0 6px 14px -6px rgba(79,70,229,.45);
}
button[kind="primary"], button[kind="primaryFormSubmit"] {
  background: linear-gradient(180deg, var(--brand-2), var(--brand)); border: 1px solid var(--brand-ink);
}

/* Upload zones */
[data-testid="stFileUploaderDropzone"] {
  min-height: 160px; padding: 1.6rem; border: 2px dashed #C7CCF8; border-radius: var(--radius);
  background: linear-gradient(180deg, #FAFAFF, #F4F5FE); transition: border-color .2s, background .2s;
  display: flex; flex-direction: column; align-items: center; justify-content: center; gap: .55rem; text-align: center;
}
[data-testid="stFileUploaderDropzone"]::before {
  content: "Drag and drop files here"; font-weight: 600; font-size: .95rem; color: var(--ink-2);
}
[data-testid="stFileUploaderDropzone"] > * { margin: 0 !important; }
[data-testid="stFileUploaderDropzoneInstructions"] { text-align: center; }
[data-testid="stFileUploaderDropzone"]:hover { border-color: var(--brand); background: #F1F2FE; }
[data-testid="stFileUploaderFile"] {
  background: var(--surface); border: 1px solid var(--line); border-radius: 12px; padding: .4rem .7rem;
  box-shadow: var(--shadow); margin-top: .4rem;
}
[data-testid="stForm"] { border: none; padding: 0; }

/* Inputs and tables */
[data-baseweb="select"] > div, [data-baseweb="input"] > div { border-radius: 10px; }
[data-testid="stDataFrame"] { border: 1px solid var(--line); border-radius: 12px; overflow: hidden; }
[data-testid="stAlert"] { border-radius: 12px; }
[data-testid="stExpander"] details { border-radius: 12px; border: 1px solid var(--line); background: var(--surface); }

/* Custom components */
.topbar { display: flex; align-items: center; justify-content: space-between; gap: 1rem; flex-wrap: wrap;
  padding: .2rem 0 1.1rem; border-bottom: 1px solid var(--line); margin-bottom: 1.4rem; }
.crumbs { font-size: .85rem; color: var(--muted); font-weight: 500; }
.crumbs b { color: var(--ink); font-weight: 600; }
.pills { display: flex; gap: .5rem; flex-wrap: wrap; }
.pill { display: inline-flex; align-items: center; gap: .35rem; font-size: .78rem; font-weight: 600; color: var(--ink-2);
  background: var(--surface); border: 1px solid var(--line); border-radius: 999px; padding: .3rem .7rem;
  box-shadow: var(--shadow); white-space: nowrap; max-width: 340px; overflow: hidden; text-overflow: ellipsis; }
.pill svg { width: 14px; height: 14px; flex: none; }
.pill.secure { color: var(--ok); background: var(--ok-soft); border-color: #ABEFC6; }

.hero { position: relative; overflow: hidden; border-radius: 18px; padding: 2.2rem 2.4rem; color: #fff;
  background: radial-gradient(1200px 400px at 85% -20%, rgba(129,140,248,.55), transparent 60%),
              linear-gradient(135deg, #1E1B4B 0%, #312E81 45%, #4338CA 100%);
  box-shadow: var(--shadow-lg); margin-bottom: 1.2rem; }
.hero h1 { color: #fff; font-size: 2rem; font-weight: 800; line-height: 1.15; margin: .4rem 0 .6rem; max-width: 720px; }
.hero p { color: #E0E7FF; font-size: 1.02rem; max-width: 680px; margin: 0 0 1.1rem; }
.hero .eyebrow { display: inline-block; font-size: .72rem; font-weight: 700; letter-spacing: .1em; text-transform: uppercase;
  color: #C7D2FE; background: rgba(255,255,255,.08); border: 1px solid rgba(255,255,255,.18); border-radius: 999px;
  padding: .25rem .65rem; }
.hero ul { display: flex; flex-wrap: wrap; gap: .5rem 1.4rem; list-style: none; padding: 0; margin: 0; }
.hero li { display: flex; align-items: center; gap: .45rem; font-size: .9rem; color: #EEF2FF; }
.hero li svg { width: 17px; height: 17px; color: #A5B4FC; }

.stats { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 1rem; margin: 0 0 1.6rem; }
.stat { background: var(--surface); border: 1px solid var(--line); border-radius: var(--radius); padding: 1.05rem 1.15rem;
  box-shadow: var(--shadow); transition: transform .2s ease, box-shadow .2s ease; position: relative; }
.stat:hover { transform: translateY(-2px); box-shadow: var(--shadow-lg); }
.stat .icon { width: 36px; height: 36px; border-radius: 10px; display: grid; place-items: center; margin-bottom: .7rem; }
.stat .icon svg { width: 19px; height: 19px; }
.stat .label { font-size: .74rem; font-weight: 600; color: var(--muted); text-transform: uppercase; letter-spacing: .06em; }
.stat .value { font-size: 1.6rem; font-weight: 750; letter-spacing: -0.02em; color: var(--ink); margin-top: .15rem; }
.stat .note { font-size: .78rem; color: var(--muted); margin-top: .1rem; }
.tip { cursor: help; }
.tip:hover::after { content: attr(data-tip); position: absolute; left: 1rem; right: 1rem; bottom: calc(100% + 6px); z-index: 10;
  background: var(--ink); color: #fff; font-size: .78rem; line-height: 1.35; padding: .5rem .65rem; border-radius: 8px;
  box-shadow: var(--shadow-lg); }

.zone-head { display: flex; align-items: center; gap: .75rem; margin: .1rem 0 .55rem; }
.badge { width: 42px; height: 42px; border-radius: 11px; display: grid; place-items: center; font-size: .72rem;
  font-weight: 800; letter-spacing: .03em; color: #fff; box-shadow: var(--shadow); flex: none; }
.badge.pdf { background: linear-gradient(160deg, #F04438, #B42318); }
.badge.xls { background: linear-gradient(160deg, #12B76A, #067647); }
.badge.book { background: linear-gradient(160deg, #6366F1, #3730A3); }
.zone-head .t { font-weight: 650; font-size: 1rem; color: var(--ink); }
.zone-head .s { font-size: .82rem; color: var(--muted); }

.result { background: var(--surface); border: 1px solid var(--line); border-left: 4px solid var(--ok);
  border-radius: var(--radius); padding: 1rem 1.2rem; box-shadow: var(--shadow); margin: .7rem 0; animation: fadeUp .3s ease both; }
.result.review { border-left-color: #F79009; }
.result .head { display: flex; justify-content: space-between; gap: 1rem; align-items: flex-start; flex-wrap: wrap; }
.result .name { font-weight: 650; color: var(--ink); font-size: 1rem; word-break: break-word; }
.result .meta { font-size: .82rem; color: var(--muted); margin-top: .15rem; }
.status { display: inline-flex; align-items: center; gap: .35rem; font-size: .74rem; font-weight: 700; letter-spacing: .04em;
  border-radius: 999px; padding: .28rem .65rem; white-space: nowrap; }
.status.ok { color: var(--ok); background: var(--ok-soft); border: 1px solid #ABEFC6; }
.status.review { color: var(--warn); background: var(--warn-soft); border: 1px solid #FEDF89; }
.status svg { width: 13px; height: 13px; }
.figs { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: .6rem; margin-top: .85rem; }
.fig { background: var(--canvas); border: 1px solid var(--line-2); border-radius: 10px; padding: .55rem .7rem; }
.fig .k { font-size: .7rem; font-weight: 600; color: var(--muted); text-transform: uppercase; letter-spacing: .05em; }
.fig .v { font-size: .98rem; font-weight: 650; color: var(--ink); margin-top: .1rem; font-variant-numeric: tabular-nums; }
.result .foot { font-size: .82rem; color: var(--ink-2); margin-top: .75rem; }
.result .foot ul { margin: .35rem 0 0 1rem; padding: 0; color: var(--warn); }

.skeleton { border-radius: var(--radius); border: 1px solid var(--line); background: var(--surface); padding: 1rem 1.2rem;
  margin: .7rem 0; box-shadow: var(--shadow); }
.skeleton .bar { height: 12px; border-radius: 6px; margin: .45rem 0;
  background: linear-gradient(90deg, #EEF0F4 0%, #F7F8FA 40%, #EEF0F4 80%); background-size: 800px 100%;
  animation: shimmer 1.2s linear infinite; }

.empty { text-align: center; background: var(--surface); border: 1px dashed #CDD3DD; border-radius: 18px;
  padding: 3rem 1.5rem; margin: 1rem 0; }
.empty .ring { width: 64px; height: 64px; margin: 0 auto 1rem; border-radius: 18px; display: grid; place-items: center;
  background: var(--brand-soft); color: var(--brand); }
.empty .ring svg { width: 30px; height: 30px; }
.empty h4 { font-size: 1.15rem; font-weight: 700; margin: 0 0 .35rem; color: var(--ink); }
.empty p { color: var(--muted); max-width: 460px; margin: 0 auto; }

.brand { display: flex; align-items: center; gap: .65rem; padding: .2rem .25rem 1rem; border-bottom: 1px solid var(--line); }
.brand .mark { width: 34px; height: 34px; border-radius: 10px; display: grid; place-items: center; color: #fff;
  background: linear-gradient(140deg, #6366F1, #312E81); box-shadow: var(--shadow); }
.brand .mark svg { width: 19px; height: 19px; }
.brand .n { font-weight: 750; font-size: 1rem; color: var(--ink); line-height: 1.1; }
.brand .s { font-size: .74rem; color: var(--muted); }
.side-foot { font-size: .74rem; color: var(--muted); border-top: 1px solid var(--line); padding-top: .8rem; margin-top: 1rem; }

@media (max-width: 900px) {
  .block-container { padding: 1rem 1rem 3rem; }
  .stats, .figs { grid-template-columns: repeat(2, minmax(0, 1fr)); }
  .hero { padding: 1.6rem 1.4rem; } .hero h1 { font-size: 1.5rem; }
}
@media (prefers-reduced-motion: reduce) { * { animation: none !important; transition: none !important; } }
</style>
""")

# Line icons (Lucide), drawn inline so they're crisp at any size.
_ICON_PATHS = {
    "file": '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/>'
            '<path d="M16 13H8M16 17H8M10 9H8"/>',
    "layers": '<path d="m12 2 10 5-10 5L2 7z"/><path d="m2 17 10 5 10-5"/><path d="m2 12 10 5 10-5"/>',
    "receipt": '<path d="M4 2v20l2-1 2 1 2-1 2 1 2-1 2 1 2-1 2 1V2l-2 1-2-1-2 1-2-1-2 1-2-1-2 1z"/>'
               '<path d="M16 8h-6a2 2 0 1 0 0 4h4a2 2 0 1 1 0 4H8M12 17.5v-11"/>',
    "shield": '<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><path d="m9 12 2 2 4-4"/>',
    "check": '<path d="M20 6 9 17l-5-5"/>',
    "alert": '<path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3z"/>'
             '<path d="M12 9v4M12 17h.01"/>',
    "lock": '<rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/>',
    "calendar": '<rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/>',
    "book": '<path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20V2H6.5A2.5 2.5 0 0 0 4 4.5z"/><path d="M4 19.5V22h16v-5"/>',
    "upload": '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m17 8-5-5-5 5M12 3v12"/>',
    "chart": '<path d="M3 3v18h18"/><path d="M7 16V9M12 16V5M17 16v-4"/>',
}


def icon(name: str) -> str:
    return (f'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" '
            f'stroke-linejoin="round" aria-hidden="true">{_ICON_PATHS[name]}</svg>')


def rand(value) -> str:
    """R 47 436.83 - South African style, a space between thousands."""
    if value is None or (isinstance(value, float) and value != value):
        return "-"
    return ("-" if value < -0.004 else "") + "R " + f"{abs(value):,.2f}".replace(",", " ")


def count(n) -> str:
    return f"{n:,}".replace(",", " ")


def esc(text) -> str:
    return escape(str(text if text is not None else ""))


def top_bar(section: str, page_name: str, pills: list[tuple[str, str, str]] = ()) -> None:
    """Breadcrumb on the left; context pills (icon, text, class) on the right."""
    chips = "".join(f'<span class="pill {cls}" title="{esc(text)}">{icon(ic)}{esc(text)}</span>'
                    for ic, text, cls in pills)
    chips += f'<span class="pill secure" title="Statements stay in this private session">{icon("lock")}Private session</span>'
    html(f'<div class="topbar"><div class="crumbs">{esc(section)} &nbsp;/&nbsp; <b>{esc(page_name)}</b></div>'
         f'<div class="pills">{chips}</div></div>')


def page_header(title: str, subtitle: str = "") -> None:
    st.header(title, anchor=False)
    if subtitle:
        st.caption(subtitle)


def section(title: str, subtitle: str = "") -> None:
    st.subheader(title, anchor=False)
    if subtitle:
        st.caption(subtitle)


_card_ids = itertools.count()


def card():
    """A card: a white, rounded, softly shadowed panel."""
    return st.container(key=f"card_{next(_card_ids)}")


def figures_row(items: list[tuple[str, str, str | None]]) -> None:
    """A row of figure cards: (label, value, help)."""
    for column, (label, value, help_text) in zip(st.columns(len(items)), items):
        with column, card():
            st.metric(label, value, help=help_text)


def money_table(df: pd.DataFrame, money: list[str], **kwargs) -> None:
    """A table with amounts in accounting format and dates as "29 Jan 2024"."""
    df = df.copy()
    config = {c: st.column_config.NumberColumn(c, format="accounting") for c in money if c in df}
    for c in money:
        if c in df:
            df[c] = pd.to_numeric(df[c], errors="coerce")  # blanks stay blank, not "None"
    for c in ("Date", "After", "At"):
        if c in df:
            df[c] = pd.to_datetime(df[c], errors="coerce")
            config[c] = st.column_config.DateColumn(c, format="D MMM YYYY")
    if "Source File" in df:
        df["Source File"] = df["Source File"].astype(str).str.replace(r"^[0-9a-f]{12}_", "", regex=True)
    st.dataframe(df, column_config=config, use_container_width=True, **kwargs)


def empty_state(icon_name: str, title: str, text: str, action: str | None = None) -> None:
    html(f'<div class="empty"><div class="ring">{icon(icon_name)}</div><h4>{esc(title)}</h4><p>{esc(text)}</p></div>')
    if action:
        st.button(action, type="primary", icon=":material/upload_file:", on_click=go_to, args=(UPLOAD,))


def go_to(page_name: str) -> None:
    st.session_state["page"] = page_name


def skeleton(cards: int) -> str:
    one = ('<div class="skeleton"><div class="bar" style="width:42%"></div><div class="bar" style="width:68%"></div>'
           '<div class="bar" style="width:90%"></div></div>')
    return one * cards


def result_card(s: dict) -> None:
    if not s["ok"]:
        st.error(f"**{s['name']}** - {s['error']}", icon=":material/error:")
        return
    approved = s["status"] == "APPROVED"
    status = (f'<span class="status ok">{icon("check")}APPROVED</span>' if approved else
              f'<span class="status review">{icon("alert")}NEEDS REVIEW</span>')
    report = s["report"]
    figs = "".join(f'<div class="fig"><div class="k">{k}</div><div class="v">{v}</div></div>' for k, v in (
        ("Opening", rand(report.get("opening_balance"))), ("Money in", rand(report.get("total_credits"))),
        ("Money out", rand(report.get("total_debits"))),
        ("Closing", rand(report.get("printed_closing") if report.get("printed_closing") is not None
                         else report.get("last_balance")))))
    added = f"{count(s['added'])} transactions added to <b>{esc(s['workbook'])}</b>"
    if s["duplicates"]:
        added += f" &middot; {count(s['duplicates'])} already in it"
    if approved:
        foot = f"{added}. Every row is confirmed by the running balance and a second, independent reading."
    else:
        issues = "".join(f"<li>{esc(p)}</li>" for p in s["problems"][:6])
        foot = (f"{added}. <b>{count(s['review'])} of {count(s['rows'])} rows</b> need checking against the "
                f"statement - highlighted in the workbook and listed under Checks."
                + (f"<ul>{issues}</ul>" if issues else ""))
    html(f'<div class="result {"" if approved else "review"}"><div class="head"><div>'
         f'<div class="name">{esc(s["name"])}</div><div class="meta">{esc(s["bank"])} &middot; {esc(s["period"])}'
         f' &middot; Account {esc(s["account"])}</div></div>{status}</div><div class="figs">{figs}</div>'
         f'<div class="foot">{foot}</div></div>')


def session_stats() -> dict:
    """What this visit has done, across its workbooks."""
    statements = rows_checked = reports = 0
    vat = 0.0
    for book in list_workbooks(settings.output_dir):
        try:
            rows = read_transactions(book)
            reviews = read_reviews(book)
        except Exception:
            continue
        statements += len(reviews) or 1
        rows_checked += len(rows)
        reports += 1
        vat += sum(abs(m["VAT on income"]) + abs(m["VAT on expenses"]) for m in vat_by_month(rows, categories))
    return {"statements": statements, "rows": rows_checked, "reports": reports, "vat": vat}


def stat_cards(stats: dict) -> None:
    cards = [
        ("file", "#EEF2FF", "#4F46E5", "Statements processed", count(stats["statements"]), "this session",
         "Statements read and added to a workbook in this visit."),
        ("layers", "#ECFDF3", "#067647", "Reports generated", count(stats["reports"] * REPORTS_PER_WORKBOOK),
         f"{count(stats['reports'])} workbook(s) x {REPORTS_PER_WORKBOOK} reports",
         "Each workbook holds the Review, Monthly Summary, Period Summary, Income Statement, Cash Flow, Category "
         "Breakdown, VAT Summary, VAT201 and per-month VAT reports."),
        ("receipt", "#FFFAEB", "#B54708", "VAT detected", rand(stats["vat"]), "calculated at 15%",
         "VAT on income and expenses marked VAT Yes - calculated, never read from the statements."),
        ("shield", "#F4F3FF", "#5925DC", "Checks performed", count(stats["rows"] * 2),
         f"{count(stats['rows'])} rows x 2 checks",
         "Every transaction is checked against the running balance and against a second, independent reading."),
    ]
    html('<div class="stats">' + "".join(
        f'<div class="stat tip" data-tip="{esc(tip)}"><div class="icon" style="background:{bg};color:{fg}">'
        f'{icon(ic)}</div><div class="label">{esc(label)}</div><div class="value">{value}</div>'
        f'<div class="note">{esc(note)}</div></div>'
        for ic, bg, fg, label, value, note, tip in cards) + "</div>")


REPORTS_PER_WORKBOOK = 9


# --- Pages ---------------------------------------------------------------------------------

def upload_page() -> None:
    top_bar("Data", UPLOAD)
    html(f'<div class="hero"><span class="eyebrow">Bank statements to financial reports</span>'
         f'<h1>Turn any South African bank statement into reconciled, audit-ready reports.</h1>'
         f'<p>Upload PDFs or spreadsheets from any bank. Every figure is kept exactly as printed and checked against '
         f'the statement&#39;s own running balance - then turned into income statements, cash flow, VAT201 '
         f'figures and monthly summaries in Excel.</p><ul>'
         f'<li>{icon("shield")}Every row checked twice</li><li>{icon("check")}Nothing guessed or corrected</li>'
         f'<li>{icon("calendar")}Monthly, quarterly &amp; yearly views</li><li>{icon("lock")}Private session</li>'
         f'</ul></div>')
    stats_slot = st.empty()
    with stats_slot:
        stat_cards(session_stats())

    opened = [Path(p) for p in st.session_state.get("opened", []) if Path(p).exists()]
    page_header(UPLOAD, "Drop in as many statements as you like. Each statement gets a workbook of its own - "
                        "statements are never combined.")
    with st.form("upload", clear_on_submit=True):
        left, right = st.columns(2, gap="large")
        with left, card():
            html(f'<div class="zone-head"><div class="badge pdf">PDF</div><div><div class="t">Bank statement PDFs</div>'
                 f'<div class="s">Digital or scanned &middot; any SA bank</div></div></div>')
            files = st.file_uploader("Bank statement PDFs - digital or scanned", type="pdf",
                                     accept_multiple_files=True, label_visibility="collapsed")
        with right, card():
            html(f'<div class="zone-head"><div class="badge xls">XLS</div><div><div class="t">Statement spreadsheets'
                 f'</div><div class="s">Excel (.xlsx) or CSV &middot; messy files welcome</div></div></div>')
            sheets = st.file_uploader(
                "Bank statement spreadsheets - Excel (.xlsx) or CSV",
                type=["xlsx", "xlsm", "csv"], accept_multiple_files=True, label_visibility="collapsed",
                help="A statement exported or typed into a spreadsheet: a header row with a date column and an "
                     "amount (or debit and credit) column, and a balance column. Titles, notes, totals and extra "
                     "columns around it are fine. Older .xls files: save them as .xlsx first.")
        options, action = st.columns([3, 1], vertical_alignment="bottom")
        with options:
            password = st.text_input(
                "PDF password (only if the statements are locked)",
                type="password",
                help="Used for this upload only. To stop typing it, add it to STATEMENT_PDF_PASSWORDS in .env.",
            )
            add_to_opened = st.checkbox(
                f"Add them to the workbook I opened ({', '.join(p.stem for p in opened)})", value=True,
            ) if opened else False
        with action:
            submitted = st.form_submit_button("Read statements", type="primary", icon=":material/play_arrow:",
                                              use_container_width=True)
    if not submitted:
        if not list_workbooks(settings.output_dir):
            st.caption("Every figure is checked against the statement's running balance; anything that can't be "
                       "confirmed is marked for review, never guessed. To add statements to a workbook you "
                       "downloaded before, open it under Continue a workbook.")
        return
    uploads = list(files or []) + list(sheets or [])
    if not uploads:
        st.warning("Choose at least one PDF or spreadsheet first.", icon=":material/upload_file:")
        return
    if not add_to_opened:
        start_fresh()

    section("Results", f"{len(uploads)} file(s) uploaded")
    bar = st.progress(0.0, text="Starting...")
    placeholder = st.empty()
    placeholder.markdown(skeleton(min(len(uploads), 3)), unsafe_allow_html=True)
    summaries = []
    for i, f in enumerate(uploads):
        bar.progress(i / len(uploads), text=f"Reading {f.name} ({i + 1} of {len(uploads)})...")

        def progress(message: str, name=f.name, i=i) -> None:
            bar.progress(i / len(uploads), text=f"{name}: {message}")

        summaries += process_upload(f.name, f.getvalue(), password, opened if add_to_opened else [], progress)
    bar.progress(1.0, text=f"Done - {len(uploads)} file(s) read.")
    with stats_slot:
        stat_cards(session_stats())  # now including what was just read
    placeholder.empty()
    for s in summaries:
        result_card(s)
    good = [s for s in summaries if s["ok"]]
    approved = sum(1 for s in good if s["status"] == "APPROVED")
    if good:
        st.toast(f"{len(good)} statement(s) read - {approved} approved, {len(good) - approved} to review.",
                 icon=":material/task_alt:")
        st.button("View the overview", type="primary", icon=":material/monitoring:", on_click=go_to,
                  args=(OVERVIEW,))


def continue_page() -> None:
    top_bar("Data", CONTINUE)
    page_header(CONTINUE, "Open a workbook you downloaded before, to add statements to it - a statement already "
                          "in it is never added twice.")
    with st.form("continue", clear_on_submit=True):
        with card():
            html(f'<div class="zone-head"><div class="badge book">XLSX</div><div><div class="t">A workbook from this '
                 f'tool</div><div class="s">The .xlsx you downloaded under Download</div></div></div>')
            uploaded = st.file_uploader("Workbook you downloaded from this page before", type="xlsx",
                                        accept_multiple_files=True, label_visibility="collapsed",
                                        help="Only a workbook this page made. Bank statements go under Upload "
                                             "statements.")
        submitted = st.form_submit_button("Open workbook", type="primary", icon=":material/folder_open:")
    if not submitted:
        return
    if not uploaded:
        st.warning("Choose a workbook first.", icon=":material/folder_open:")
        return
    start_fresh()
    opened = []
    for w in uploaded:
        target = restore_workbook(w.getvalue(), settings.output_dir, w.name)
        if target is None:
            st.error(f"**{w.name}** isn't a workbook downloaded from this page, so it was not opened.",
                     icon=":material/block:")
            continue
        opened.append(str(target))
        st.session_state["last_workbook"] = str(target)
        st.success(f"Opened **{target.stem}** ({count(len(read_transactions(target)))} transactions). Go to "
                   "**Upload statements** to add statements to it.", icon=":material/check_circle:")
        st.toast(f"Opened {target.stem}", icon=":material/folder_open:")
    st.session_state["opened"] = opened
    if opened:
        st.button("Add statements to it", type="primary", icon=":material/upload_file:", on_click=go_to,
                  args=(UPLOAD,))


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


def money_in_out_chart(df: pd.DataFrame):
    """Money in and out side by side per month, in date order."""
    monthly = df.groupby("Month")[["Credit", "Debit"]].sum().reset_index()
    order = [month_label(m) for m in monthly["Month"]]
    long = monthly.melt(id_vars="Month", value_vars=["Credit", "Debit"], var_name="Flow", value_name="Rand")
    long["Flow"] = long["Flow"].map({"Credit": "Money in", "Debit": "Money out"})
    long["Month"] = [month_label(m) for m in long["Month"]]
    return (alt.Chart(long, height=280)
            .mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3)
            .encode(x=alt.X("Month:N", sort=order, title=None, axis=alt.Axis(labelAngle=-45)),
                    xOffset=alt.XOffset("Flow:N"),
                    y=alt.Y("Rand:Q", title=None, axis=alt.Axis(format=",.0f", grid=True)),
                    color=alt.Color("Flow:N", scale=alt.Scale(range=["#12B76A", "#F04438"]),
                                    legend=alt.Legend(orient="bottom", title=None)),
                    tooltip=["Month", "Flow", alt.Tooltip("Rand:Q", format=",.2f")])
            .configure_view(strokeWidth=0).configure_axis(labelColor="#5B6779", gridColor="#EEF0F4",
                                                          domainColor="#E4E7EC"))


def overview_page(rows: list[dict], period: Period) -> None:
    page_header(OVERVIEW, f"{period.label} - from {count(len(rows))} transactions, exactly as on the statements.")
    df = _frame(rows)
    income = df[df["Type"] == "income"]
    expenses = df[df["Type"] == "expense"]
    total_income = income["Credit"].sum() - income["Debit"].sum()
    total_expenses = expenses["Debit"].sum() - expenses["Credit"].sum()
    opening, closing = _opening_closing(rows)
    vat = sum(m["VAT payable"] for m in vat_by_month(rows, categories))
    figures_row([("Opening balance", rand(opening), "Before the period's first transaction"),
                 ("Money in", rand(df["Credit"].sum()), "Everything credited to the account"),
                 ("Money out", rand(df["Debit"].sum()), "Everything debited from the account"),
                 ("Closing balance", rand(closing), "After the period's last transaction")])
    figures_row([("Income", rand(total_income), "Money in, categorised as income"),
                 ("Expenses", rand(total_expenses), "Money out, categorised as expenses"),
                 ("Net profit", rand(total_income - total_expenses), "Income minus expenses; transfers and "
                                                                        "cash withdrawals aren't expenses"),
                 ("VAT payable", rand(vat), f"Calculated at {VAT_RATE:.0%} - not read from the statements")])

    left, right = st.columns(2)
    with left, card():
        st.markdown("**Money in and out by month**")
        st.altair_chart(money_in_out_chart(df), use_container_width=True)
    with right, card():
        st.markdown("**Spending by category**")
        spend = defaultdict(float)
        for _, r in expenses.iterrows():
            spend[r["Category"]] += r["Debit"] - r["Credit"]
        if spend:
            st.bar_chart(pd.Series(spend, name="Rand").sort_values(ascending=False), horizontal=True,
                         color="#4F46E5", height=280)
        else:
            st.caption("No expenses in this period.")

    # Income statement: each category's net per month (income as money in,
    # everything else as money out); transfers and drawings apart, outside profit.
    df["Net"] = df["Credit"] - df["Debit"]
    df.loc[df["Type"] != "income", "Net"] *= -1
    for title, kinds, note in (
            ("Income statement", ("income", "expense"), "Net per category and month, cash basis."),
            ("Not in profit", ("transfer", "drawings"), "Transfers, cash withdrawals and drawings - money out, net.")):
        part = df[df["Type"].isin(kinds)]
        if part.empty:
            continue
        section(title, note)
        table = part.pivot_table(index=["Type", "Category"], columns="Month", values="Net", aggfunc="sum",
                                 fill_value=0.0)
        table.columns = [month_label(m) for m in table.columns]
        table["Total"] = table.sum(axis=1)
        money_table(table.round(2), list(table.columns))
    st.caption("Categories come from config/categories.yaml, your own edits in the workbook, and suggestions "
               "learnt from them - see Checks.")


def monthly_page(rows: list[dict], period: Period) -> None:
    page_header(MONTHLY, f"Each month of {period.label} - the workbook's Monthly Summary and Period Summary sheets "
                         "have the same figures.")
    table = monthly_table(rows)
    if table.empty:
        empty_state("calendar", "No dated transactions", "This period has no transactions with a readable date.")
        return
    totals = {c: table[c].sum() for c in ("Money in", "Money out", "Income", "Expenses", "Net profit",
                                          "VAT payable", "Transactions", "To review")}
    totals.update({"Month": "Total", "Opening balance": table["Opening balance"].iloc[0],
                   "Closing balance": table["Closing balance"].iloc[-1]})
    shown = pd.concat([table, pd.DataFrame([totals])], ignore_index=True)
    money_table(shown.round(2), MONEY_COLUMNS, hide_index=True)

    section("One month in detail", "Pick a month to see its figures, categories and transactions.")
    month = st.selectbox("Month", list(table["Month"]), index=len(table) - 1)
    key = next(m for m in period.months if month_label(m) == month)
    part = [r for r in rows if r.get("Month") == key]
    line = table[table["Month"] == month].iloc[0]
    figures_row([("Money in", rand(line["Money in"]), None), ("Money out", rand(line["Money out"]), None),
                 ("Net profit", rand(line["Net profit"]), None),
                 ("Closing balance", rand(line["Closing balance"]), None)])
    df = _frame(part)
    by_category = df.groupby(["Type", "Category"])[["Credit", "Debit"]].sum().rename(
        columns={"Credit": "Money in", "Debit": "Money out"}).reset_index()
    money_table(by_category.round(2), ["Money in", "Money out"], hide_index=True)
    with st.expander(f"All {count(len(part))} transactions of {month}", icon=":material/list:"):
        money_table(df[[c for c in ("Date", "Description", "Category", "Debit", "Credit", "Balance", "Status")
                        if c in df]], ["Debit", "Credit", "Balance"], hide_index=True)


def vat_page(rows: list[dict], all_rows: list[dict], period: Period) -> None:
    page_header(VAT, f"{period.label}. VAT is calculated at {VAT_RATE:.0%} on transactions marked VAT Yes - never "
                     "read from the statements.")
    if not st.session_state.get("vat_registered", True):
        st.info("**Not VAT-registered** - no VAT is charged on income or claimed on expenses, so every VAT figure is "
                "R 0.00. If the business is VAT-registered, switch on **VAT-registered business** in the side menu.",
                icon=":material/info:")
    months = vat_by_month(rows, categories)
    if not months:
        empty_state("receipt", "No VAT in this period", "There are no income or expense transactions in it.")
        return
    table = pd.DataFrame(months)
    table["Month"] = [month_label(m) for m in table["Month"]]
    table = table.set_index("Month")[["Income", "VAT on income", "Expenses", "VAT on expenses", "VAT payable"]]
    totals = table.sum()  # added up unrounded, as the workbook does
    figures_row([("VAT on income (output)", rand(totals["VAT on income"]), "Field 4 / 13 of the VAT201"),
                 ("VAT on expenses (input)", rand(totals["VAT on expenses"]), "Field 15 / 19 - claim only with a "
                                                                              "valid tax invoice"),
                 ("VAT payable", rand(totals["VAT payable"]), "Field 20 - negative means a refund")])
    table.loc["Total"] = totals
    money_table(table.round(2), list(table.columns))

    section("VAT201 return", "The SARS VAT201 fields bank statements can support. Capital goods belong in field 14; "
                             "income without VAT goes in field 2 (zero-rated) or 3 (exempt).")
    all_months = [r.get("Month") for r in all_rows]
    with card():
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

    with st.expander("Transactions and their VAT setting", icon=":material/list:"):
        df = _frame(rows)
        money_table(df[df["Type"].isin(["income", "expense"])][
            [c for c in ("Date", "Description", "Category", "VAT", "Debit", "Credit") if c in df]],
            ["Debit", "Credit"], hide_index=True)
    st.caption("Change a transaction's VAT Yes/No in the workbook's Transactions sheet. The workbook has the same "
               "figures per month, in VAT Summary and VAT201.")


def checks_page(rows: list[dict], selected: Path) -> None:
    page_header(CHECKS, "How each statement reconciled, and anything that needs a person to look at it - for the "
                        "whole workbook.")
    df = _frame(rows)
    reviews = read_reviews(selected)
    to_review = df[df["Status"] == "REVIEW_REQUIRED"] if "Status" in df else df.iloc[0:0]
    gaps = workbook_gaps(rows)
    figures_row([("Transactions", count(len(df)), None),
                 ("Confirmed", count(len(df) - len(to_review)),
                  "Confirmed by the running balance and by a second, independent reading"),
                 ("To review", count(len(to_review)), "Kept exactly as printed and marked, with the reason"),
                 ("Balance gaps", count(len(gaps)), "Places where the balance doesn't carry on from one "
                                                    "transaction to the next")])
    if reviews:
        section("Statements")
        money_table(pd.DataFrame(reviews)[[c for c in ("Source File", "Statement Period", "Status",
                                                        "Rows Needing Review", "Issues") if c in reviews[0]]],
                    [], hide_index=True)

    section("Balance check", "The running balance must carry on unbroken from one transaction to the next, across "
                             "every statement in the workbook.")
    if gaps:
        st.error(f"**The balances don't join up in {len(gaps)} place(s)** - usually a missing statement for that "
                 "period, or a transaction edited in Excel. Totals for those months are incomplete.",
                 icon=":material/link_off:")
        money_table(pd.DataFrame({
            "After": g.after_date, "At": g.at_date, "Transaction": g.description,
            "Expected balance": g.expected, "Statement balance": g.actual,
            "Missing amount": round(g.actual - g.expected, 2),
        } for g in gaps), ["Expected balance", "Statement balance", "Missing amount"], hide_index=True)
    else:
        st.success("Every transaction follows on from the one before - no gaps or misreads.",
                   icon=":material/check_circle:")

    section("Rows to review")
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

    section("Categories")
    uncategorised = df[df["Category"].str.contains("uncategorised", na=False)]
    if len(uncategorised):
        st.warning(f"{len(uncategorised)} transaction(s) are uncategorised. Type a category into the Category column "
                   "in Excel, or add a rule in config/categories.yaml and click **Re-apply category rules**.",
                   icon=":material/label:")
        money_table(uncategorised[["Date", "Description", "Debit", "Credit"]], ["Debit", "Credit"], hide_index=True)
    suggested = (df[df["Category Source"].astype(str).str.startswith("Suggested")] if "Category Source" in df
                 else df.iloc[0:0])
    if len(suggested):
        st.info(f"{len(suggested)} transaction(s) have a **suggested** category, learnt from similar transactions "
                "whose category you typed. Check them in the Category column in Excel: type over any that are "
                "wrong, and later suggestions learn from it.", icon=":material/auto_awesome:")
        money_table(suggested[["Date", "Description", "Category", "Category Source", "Debit", "Credit"]],
                    ["Debit", "Credit"], hide_index=True)
    if not len(uncategorised) and not len(suggested):
        st.success("Every transaction has a category from a rule or from you.", icon=":material/check_circle:")
    if st.button("Re-apply category rules", icon=":material/refresh:"):
        try:
            append_transactions(selected, [], categories)
        except WorkbookLockedError as exc:
            st.error(str(exc))
        else:
            st.session_state["toast"] = "Category rules re-applied"
            st.rerun()


def download_page(rows: list[dict], selected: Path) -> None:
    page_header(DOWNLOAD, "Excel workbooks with every report. Each cell holds its number (for phones and previews) "
                          "and its formula (so Excel recalculates your edits).")
    books = list_workbooks(settings.output_dir)
    if len(books) > 1:
        bundle = io.BytesIO()
        with zipfile.ZipFile(bundle, "w", zipfile.ZIP_DEFLATED) as z:
            for book in books:
                z.write(book, book.name)
        with card():
            html(f'<div class="zone-head"><div class="badge book">ZIP</div><div><div class="t">All {len(books)} '
                 f'workbooks</div><div class="s">One per statement, in one .zip</div></div></div>')
            st.download_button(f"Download all {len(books)} workbooks (.zip)", data=bundle.getvalue(),
                               file_name="statement workbooks.zip", mime="application/zip", type="primary",
                               icon=":material/folder_zip:",
                               on_click=lambda: st.toast("Download started", icon=":material/download:"))
    periods_in = sorted({str(r.get("Statement Period")) for r in rows if r.get("Statement Period")})
    with card():
        html(f'<div class="zone-head"><div class="badge xls">XLSX</div><div><div class="t">{esc(selected.stem)}</div>'
             f'<div class="s">{count(len(rows))} transactions &middot; {len(periods_in)} statement period(s): '
             f'{esc(", ".join(periods_in))}</div></div></div>')
        st.download_button(
            "Download this workbook",
            data=selected.read_bytes(),
            file_name=selected.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary" if len(books) <= 1 else "secondary",
            icon=":material/download:",
            on_click=lambda: st.toast("Download started", icon=":material/download:"),
        )
    section("What's inside", "Review, Monthly Summary, Period Summary (months, quarters and financial years), Income "
                             "Statement, Cash Flow, Category Breakdown, VAT Summary, VAT201, one VAT sheet per "
                             "month, Transactions and Categories. Keep it: to add next month's statements, open it "
                             "under Continue a workbook.")


# --- Shell: sidebar navigation and routing ------------------------------------------------------

UPLOAD, CONTINUE, OVERVIEW, MONTHLY, VAT, CHECKS, DOWNLOAD = (
    "Upload statements", "Continue a workbook", "Overview", "Monthly", "VAT", "Checks", "Download")
ICONS = {UPLOAD: ":material/upload_file:", CONTINUE: ":material/folder_open:", OVERVIEW: ":material/monitoring:",
         MONTHLY: ":material/calendar_month:", VAT: ":material/receipt_long:", CHECKS: ":material/fact_check:",
         DOWNLOAD: ":material/download:"}
GROUP = {UPLOAD: "Data", CONTINUE: "Data", OVERVIEW: "Reports", MONTHLY: "Reports", VAT: "Reports",
         CHECKS: "Review & export", DOWNLOAD: "Review & export"}
CUSTOM = "Custom range of months"

def _vat_registration_changed() -> None:
    """Every workbook's VAT Yes/No worked out afresh for the new setting."""
    now = for_vat_registration(all_categories, st.session_state["vat_registered"])
    for book in list_workbooks(settings.output_dir):
        try:
            append_transactions(book, [], now, reset_vat=True)
        except WorkbookLockedError:
            continue
    st.session_state["toast"] = ("VAT-registered: VAT is included in payments received" if
                                 st.session_state["vat_registered"] else "Not VAT-registered: no VAT charged or claimed")


if toast := st.session_state.pop("toast", None):
    st.toast(toast, icon=":material/task_alt:")

with st.sidebar:
    html(f'<div class="brand"><div class="mark">{icon("book")}</div><div><div class="n">Bank Statement Tool</div>'
         f'<div class="s">Statements to financial reports</div></div></div>')
    page = st.radio("Menu", list(ICONS), key="page", format_func=lambda p: f"{ICONS[p]}  {p}",
                    label_visibility="collapsed")
    st.toggle("VAT-registered business", value=True, key="vat_registered", on_change=_vat_registration_changed,
              help="On: payments received include 15% VAT, and VAT on expenses is claimed. Off: the business "
                   "isn't VAT-registered, so no VAT is charged or claimed. Changing it updates every workbook.")
categories = for_vat_registration(all_categories, st.session_state["vat_registered"])
account_slot = st.sidebar.container()  # filled once this run's uploads are in


def choose_account() -> Path | None:
    books = list_workbooks(settings.output_dir)
    with account_slot:
        html('<div class="side-foot"></div>')
        if not books:
            st.caption("No workbook yet - upload statements to start one.")
            return None
        last = st.session_state.get("last_workbook")
        return st.selectbox("Workbook", books, index=next((i for i, b in enumerate(books) if str(b) == last), 0),
                            format_func=lambda p: p.stem, help="One workbook per statement uploaded")


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
                             index=labels.index(remembered) if remembered in labels else 0,
                             help="Financial years run March to February. Overview, Monthly and VAT follow it.")
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
            top_bar(GROUP[page], page)
            st.warning("The workbook is open in Excel. Close it there and refresh this page.")
            st.stop()
    if not rows:
        top_bar(GROUP[page], page)
        page_header(page)
        empty_state("upload", "No statements yet",
                    "No transactions yet - upload bank statements (PDF, Excel or CSV) and every report fills in "
                    "here.", action="Upload statements")
        st.stop()
    if page in (OVERVIEW, MONTHLY, VAT):
        period = choose_period(rows)
        top_bar(GROUP[page], page, [("book", selected.stem, ""), ("calendar", period.label, "")])
        in_period = rows if period.kind == "all" else [r for r in rows if r.get("Month") in period.months]
        if not in_period:
            page_header(page)
            empty_state("calendar", "Nothing in this period", "Choose another reporting period in the side menu.")
            st.stop()
        if page == OVERVIEW:
            overview_page(in_period, period)
        elif page == MONTHLY:
            monthly_page(in_period, period)
        else:
            vat_page(in_period, rows, period)
    else:
        top_bar(GROUP[page], page, [("book", selected.stem, "")])
        if page == CHECKS:
            checks_page(rows, selected)
        else:
            download_page(rows, selected)
