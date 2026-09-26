# statement-tool

Turns bank statement PDFs into an Excel workbook of transactions plus
financial reports (monthly summary, income statement, cash flow, category
breakdown). Built incrementally:

**Live app:** https://bank-statement-tool-za.streamlit.app/ (password
protected - see "Running it online" below)

- **Phase 1 (done):** folder of PDFs -> combined `.xlsx`.
- **Phase 2 (done):** Gmail search/download layered on top of the same
  parser, using the Gmail API with OAuth (optional - uploading is the main
  way in).
- **Phase 3 (done):** upload page, transaction categories, and the
  financial report sheets.

## Quick start: upload page

Double-click **`run_app.bat`** (or run
`.venv\Scripts\streamlit run src\statement_tool\app.py`). A page opens in
your browser at http://localhost:8501:

1. Drop in one or more statement PDFs. If they're password-protected, type
   the password too.
2. Click **Add to workbook**. Each statement is read, its transactions are
   categorised and added to its account's workbook in `output/` (e.g.
   `output/FTW Properties (10237421516).xlsx`), and the page
   shows income, expenses, profit and charts.
3. Click **Download Excel workbook** for the full reports.

Uploading the same statement twice, or statements that overlap (a 6-month
statement plus the monthly ones), never double-counts: each transaction is
recognised by its own date, description, amount and balance.

## Setup

```powershell
cd bank-statement-tool
python -m venv .venv
.venv\Scripts\pip install -e .
copy .env.example .env
```

### OCR fallback (scanned statements) - optional but recommended

The primary path (`pdfplumber`) handles digitally-generated PDFs, which is
most bank statements. If a statement turns out to be a scan/image, the tool
falls back to OCR, which needs two external programs on PATH (or pointed to
via `.env`):

1. **Tesseract OCR** - https://github.com/UB-Mannheim/tesseract/wiki (Windows installer)
2. **Poppler for Windows** - https://github.com/oschwartz10612/poppler-windows/releases

After installing, either add both `...\tesseract.exe` and
`...\poppler\Library\bin` to your PATH, or set `TESSERACT_CMD` and
`POPPLER_PATH` in `.env`. Without these, statements that need OCR will be
reported as failures in the run summary (not silently dropped) rather than
crashing the run.

## Configuration

Three YAML files, all meant to be hand-edited - no code changes needed to
recategorise spending, add a client, or add a straightforward new bank:

- **`config/categories.yaml`** - the reporting categories. Each has a type
  (`income`, `expense`, `transfer` between own accounts, or `drawings` for
  personal money out) and the description text that puts a transaction in
  it. Unmatched transactions land in "Other income/expenses
  (uncategorised)"; add a rule and click **Re-apply category rules** on the
  upload page (or upload anything) to file them. You can also type a
  category straight into a transaction's Category cell in Excel - hand
  edits are kept on every later run.

- **`config/clients.yaml`** - maps a client name to the sender/subject/
  filename/statement-text substrings that identify their statements. If a
  statement matches nothing, the tool asks for the client name interactively
  (or labels it `UNMAPPED_CLIENT` and flags it for review in non-interactive
  runs).
- **`config/banks.yaml`** - per-bank detection keywords, statement-period
  regex, and column-header aliases (date/description/debit/credit/balance,
  or a single signed `amount` column for banks like Capitec). A `generic`
  fallback with widened aliases is used when no bank's `detect` keywords
  match, and such statements are always flagged in the summary so you know
  to double check them.

Edit these directly - see the comments in each file for the exact format.

### Password-protected statements

Many banks (Standard Bank included) password-protect the statements they
email. List the password(s) in `.env`, comma-separated - each is tried in
turn until one opens the file:

```
STATEMENT_PDF_PASSWORDS=1234567890,9876543210
```

A statement that none of them open is reported as a failure in the run
summary, naming this setting.

## Running (phase 1: local folder)

Drop PDFs into `data/incoming_pdfs/` (or point `--folder` elsewhere), then:

```powershell
.venv\Scripts\python -m statement_tool.cli process-folder
```

- Already-processed PDFs (tracked by file hash in
  `data/processed/processed.db`) are skipped automatically on the next run.
- New transactions are appended to their account's workbook in `output/` on a
  single `Transactions` sheet, formatted as an Excel Table with a `TOTAL`
  row driven by real `=SUM(...)` formulas (they recalculate if you edit a
  cell), plus the report sheets below, all rebuilt each run from whatever
  is currently in `Transactions`.
- A run summary prints how many PDFs were scanned/skipped/processed, how
  many transactions were added, and lists anything that failed to parse or
  needs a manual spot-check (OCR'd statements, unmapped clients, unparsable
  dates), with the reason - nothing is ever silently dropped.

Useful flags:

- `--folder PATH` - use a different source folder.
- `--reprocess` - reprocess files even if already recorded (useful while
  tuning a new bank's column config).
- `--no-interactive` - never prompt for client mapping (for scripted runs);
  unmatched statements are labeled `UNMAPPED_CLIENT` instead.

## Running it online (Streamlit Community Cloud)

The same page can be hosted so it opens from any phone or computer:

1. Sign in at https://share.streamlit.io with the GitHub account that owns
   this repo, and allow it access to the repo.
2. **Create app** -> deploy from GitHub, then choose this
   repo, branch `main`, main file `src/statement_tool/app.py`. Under
   **Advanced settings**, pick Python 3.12 and paste into **Secrets**:
   ```
   APP_PASSWORD = "a long password of your choosing"
   ```
3. Deploy. Every push to `main` updates the app automatically.
4. In the app's **Settings -> Sharing**, keep it private and invite only
   the people who should use it.

Online, the page won't open without `APP_PASSWORD`, and it keeps nothing
between visits: each visit works in its own temporary folder, which is
gone when you leave. To add statements to an existing workbook, upload
that workbook along with the new PDFs, then **download the updated
workbook** before closing the page. `packages.txt` installs the OCR tools
there, so scanned statements work online too.

## Safety checks

Every statement is checked before anything is written, using the figures
the bank prints itself:

- **Running balance:** each transaction's balance must equal the previous
  balance plus its credit minus its debit. A misread amount, a debit read
  as a credit, or a skipped row breaks this.
- **Statement totals:** where the statement prints its own totals (Standard
  Bank's "Payments"/"Deposits" summary, or any "Closing balance" line), the
  transactions read must add up to them exactly. This catches a missing
  first or last row.
- **Dates and unreadable lines:** a date that can't be read, or a line that
  looks like a transaction but doesn't match the expected shape, is flagged
  rather than dropped.

A statement that fails any check is **not added** - the upload page (or the
command-line run summary) says exactly what didn't add up. Once you've
compared it with the PDF yourself, tick **Add even if checks fail** (or use
`--allow-problems`) to add it anyway.

The workbook itself is protected too:

- The upload page checks the balance chain across the **whole workbook**
  every time it opens, so a missing month between two statements shows up.
- The workbook is saved to a temporary file and swapped in, so a crash
  can't leave it half-written, and the previous 30 versions are kept in
  `output/backups/`.
- If the workbook is open in Excel, nothing is changed and you're told to
  close it.
- The month sheets find each transaction by its key, so sorting or
  filtering the Transactions sheet in Excel can't mix up their figures.

What the checks can't know: whether a transaction is in the right
**category** or has the right **VAT** setting - those are judgement calls,
so glance at the "uncategorised" list on the upload page and at the VAT
column.

## The workbook

Each bank account gets its **own workbook**, named after the client in
`config/clients.yaml` and the account number found on the statement (e.g.
`FTW Properties (10237421516).xlsx`, or `Account 1234567890.xlsx` for an
account with no client rule). Statements from different accounts are never
mixed, so every balance and total is about one account. On the upload page,
pick the account from the list to see its overview and download it.

| Sheet | What it shows |
| --- | --- |
| Monthly Summary | Per month: opening balance, money in, money out, net, closing balance, and a check that should be 0 |
| Income Statement | Income and expense categories by month, net profit; transfers and drawings listed separately, outside profit |
| Cash Flow | Opening balance, cash from operations, transfers, drawings, closing balance, reconciled to the bank statement |
| Category Breakdown | Count, money in/out and share of expenses per category, with a chart |
| VAT Summary | Per month: income, expenses and difference, each incl. VAT, excl. VAT and the 15% VAT, plus VAT payable |
| Jul 2026, Aug 2026, ... | One sheet per month: INCOME and EXPENSES listed separately (with references like JUL26R01 / JUL26P01), each amount split into excl. VAT and 15% VAT, with totals and a month summary |
| Transactions | Every transaction, in date order, with its Month and Category |
| Categories | The rules from `config/categories.yaml`, for reference |

Whether a transaction includes VAT is its **VAT** (Yes/No) cell on the
Transactions sheet, defaulted from its category (`vat:` in
`config/categories.yaml`); change it there and the month sheets update.
Transfers and personal payments appear in neither list.

Totals are live `SUMIFS` formulas, so correcting a transaction or its
category in Excel updates every report. Opening and closing balances come
from the running balance printed on the statements. A non-zero value in a
"Check" or "Difference" row means a month is missing a statement or a
transaction was misread.

## Running (phase 2: Gmail)

### One-time Google Cloud setup

1. Go to https://console.cloud.google.com/ and create a project (or reuse one).
2. **APIs & Services > Library** - enable the **Gmail API**.
3. **APIs & Services > OAuth consent screen** - choose **External**, fill in
   the required fields, and add your own Gmail address as a **test user**
   (this keeps the app private to you without needing Google's review).
4. **APIs & Services > Credentials > Create Credentials > OAuth client ID**
   - application type **Desktop app**. Download the resulting JSON.
5. Save that file as `credentials.json` in the project root (or anywhere,
   and point `GMAIL_CLIENT_SECRET_FILE` in `.env` at it).

### Running it

```powershell
.venv\Scripts\python -m statement_tool.cli fetch-email
```

The first run opens a browser window for you to sign in and grant
**read-only** Gmail access (the tool never sends, labels, or deletes
anything); it then caches the token to `token.json` (path configurable via
`GMAIL_TOKEN_FILE`) so every run after that is silent - no repeated login.

What it does:

1. Builds a Gmail search query: any email with a PDF attachment in the last
   `GMAIL_LOOKBACK_DAYS` (default 30) days, narrowed to senders/subjects
   mentioned in `config/clients.yaml` once that file has real entries (an
   empty/example config just searches all PDF-attachment mail).
2. For each match, downloads PDF attachments not already recorded (by
   message ID + attachment filename, and independently by content hash -
   re-running is always safe), saving them to `data/email_downloads/`.
3. Runs each one through the exact same parser/writer as `process-folder`,
   using the email's sender/subject as extra signal for client matching.

Flags: `--days N` (override the lookback window), `--reprocess`,
`--no-interactive` - same meaning as in `process-folder`.

## Adding a new bank's layout

Usually not needed: a statement from a bank without its own entry is read
by trying every row shape the tool knows (Standard Bank, Capitec, FNB and
plain layouts) and keeping the reading whose numbers match the statement's
own balances and totals. With every bank's settings hidden, all the real
statements tested so far read identically. A new entry helps when that
fails, and adds the bank's printed totals as an extra check.


Copy a block in `config/banks.yaml`, rename the key, and adjust:

- `detect` - a few substrings unique to that bank's statement header/footer.
- `columns` - the header text aliases pdfplumber will see in that bank's
  table headers.
- `amount_style` - `split` (separate debit/credit columns, most SA banks)
  or `signed` (one amount column, negative = debit).

No Python changes needed for a normal text-based statement with a
recognizable header row. If pdfplumber can't split a statement's table into
columns, a line-based fallback (`line_extract.py`) reads rows shaped
`date  description  amount  balance` straight from the text, using the
running balance to confirm which side (debit/credit) each amount is on -
this is what handles Standard Bank's 6-month statement. Layouts that fit
neither are a case for a new heuristic in `text_extract.py`/`column_map.py`.

## Tests

```powershell
.venv\Scripts\python tests\fixtures\generate_fixtures.py   # regenerate synthetic sample PDFs
.venv\Scripts\python -m pytest tests\ -v
```

The fixtures are synthetic (not real statements) - they exist to pin down
the extraction/normalization logic (column mapping, amount parsing, date
parsing, dedup) while waiting on real samples to tune the bank configs
against. Swap in real (or redacted) statements under `data/incoming_pdfs/`
and run with `--reprocess` to compare output by hand before trusting it.

## Project layout

```
config/
  categories.yaml       # reporting categories + description matching (edit me)
  clients.yaml          # client name <- sender/subject/filename/text matching (edit me)
  banks.yaml             # per-bank column layout + detection (edit me)
src/statement_tool/
  cli.py                 # command-line entry point
  config.py               # loads the YAML + .env settings above
  models.py                # Transaction / StatementResult dataclasses
  store.py                  # SQLite "already processed" tracking
  app.py                    # upload page (Streamlit) - started by run_app.bat
  categorize.py              # assigns categories from config/categories.yaml
  excel_writer.py            # workbook: transactions + report sheets
  gmail_client.py            # OAuth login, search, attachment download
  extract/
    bank_detect.py            # bank + client + statement-period detection
    column_map.py              # header row -> logical column mapping
    text_extract.py             # pdfplumber-based table extraction (primary)
    line_extract.py              # line-by-line fallback when table columns can't be split
    ocr_extract.py                # pytesseract fallback for scanned PDFs
    amounts.py, dates.py           # shared value parsing
    parser.py                       # orchestrates the above per PDF
data/incoming_pdfs/       # drop PDFs here for phase 1
data/uploads/              # PDFs added through the upload page
data/email_downloads/      # PDFs fetched via Gmail land here
data/processed/            # processed.db (SQLite dedup record)
output/                      # one workbook per bank account (+ backups/)
tests/
```

## Credentials

Phase 1 needs no credentials. Phase 2 stores the Gmail OAuth client secret
and token cache as local files referenced from `.env` (`credentials.json`,
`token.json` by default) - both are already in `.gitignore`, along with
`.env` itself and the generated workbook/database, so none of it reaches
source control.
