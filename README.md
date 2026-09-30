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

The side menu has a page for each job:

- **Upload statements** - drop in one or more statement PDFs (and the
  password, if they're locked) and click **Read statements**. Each upload
  makes a **new workbook** (one per bank account) from just the
  statements uploaded then; statements from earlier uploads never carry
  over.
- **Continue a workbook** - open a workbook you downloaded before; the
  statements you upload next are added to it.
- **Financials** - income, expenses, profit, charts, and the income
  statement by category and month.
- **VAT** - VAT on income and expenses per month, and VAT payable
  (calculated at 15%, never read from the statements).
- **Checks** - each statement's status, the balance check, rows needing
  review, and uncategorised or suggested categories.
- **Download** - the Excel workbook with all the report sheets. Keep it:
  it's how you add next month's statements (Continue a workbook).

The page works in a temporary folder of its own; it doesn't use the
`output/` folder (the command-line tool below still does).

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

### Scanned statements (OCR)

Scanned (photo / image-only) statements are read with **Tesseract OCR**:

- **Windows:** `winget install UB-Mannheim.TesseractOCR` (found automatically
  in its default folder; otherwise set `TESSERACT_CMD` in `.env`).
- **Online (Streamlit Cloud):** installed automatically from `packages.txt`.

Each page is turned the right way up (scanners often feed pages upside
down), read twice independently (two OCR passes at different resolutions),
and cleaned only of what scanning adds (table lines read as `|`, stray
marks) - amounts are never changed, not even a decimal comma. A PDF holding
several statements (e.g. a year of monthly statements scanned together) is
split per statement, and each is checked on its own. Any row the two OCR
passes disagree on, or that doesn't add up, is marked REVIEW_REQUIRED.
Better scans (straight, 300 dpi, not faded) read best.

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

  **Learnt categories.** What no rule matches gets a *suggested* category
  when it clearly resembles transactions whose category you typed, with
  money going the same way, and has their most distinctive word, usually
  the shop or person paid. For example, give one "SASOL THERESA" purchase
  the category Fuel, and the other Sasol Theresa purchases get Fuel
  suggested; how a payment was made ("DEBIT CARD PURCHASE", "PAYSHAP PAY
  BY PROXY") and its month never count. It learns only from your own
  categories, not the rules' (the rules already cover what they match). It
  learns from the workbook itself each time (nothing is stored or sent
  anywhere), so every category you type teaches it. Suggestions only set the Category, never an
  amount, and the **Category Source** column says which set each one:
  `Rule`, `Set by you`, or `Suggested - like N transactions you
  categorised, e.g. "..."`. Type over a wrong suggestion and it's kept as yours. A rule
  added later replaces a suggestion, but never a category you typed.

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
gone when you leave. Each upload makes a new workbook from just the files
uploaded that time - statements added earlier in the visit aren't carried
over. To add statements to an existing workbook, open that workbook under
**Continue a workbook**, upload the new PDFs, then **download the updated
workbook** before closing the page. `packages.txt` installs the OCR tools
there, so scanned statements work online too.

## How statements are checked

The reader works on **any bank's layout**, without per-bank setup: it uses
the arithmetic every statement has - each balance is the previous balance
plus or minus the amount - to tell which number is the amount, which the
balance, and whether money came in or went out. It is tested on hundreds of
made-up statements in layouts no bank config describes
(`tests/fixtures/synthetic.py`).

Figures are read as written, `1,234.56` or with a decimal comma (`1 234,56`,
`1.234,56`, as ABSA prints them): each statement is read in the style most
of its figures use, so `12,50` is never taken for a thousands number. An
entry printed over two lines with its date once - a payment, then its
`Service Fee` and the balance on the next line - is read as one dated
entry; any other line without a printed date stays undated and is marked.
Where a statement totals its fees apart from its debits (`Total service
fees`), the debits read must equal the two together.

Rules it follows:

- **Nothing is guessed, changed or filled in.** Amounts, dates,
  descriptions and balances are kept exactly as printed. A value the
  document doesn't establish is left empty - an unreadable date stays blank;
  an amount whose direction isn't shown goes in the **In/Out Not Shown**
  column, in no total - and the row is marked **REVIEW_REQUIRED** with the
  reason.
- **Two independent extractions.** Digital PDFs are read by two separate
  PDF text engines (pdfminer and MuPDF), scans by two OCR passes; any row
  they disagree on is REVIEW_REQUIRED.
- **Validation:** every row needs a date and an amount; each balance must
  follow from the previous one; opening + credits - debits = closing; the
  statement's printed totals and closing balance must match; duplicates,
  impossible dates and dates out of order are flagged. If two different
  readings of a row both add up, it's flagged as ambiguous.
- **Status and evidence:** each row gets a Status (APPROVED /
  REVIEW_REQUIRED), a Confidence (1.00 verified; 0.97 verified from a scan;
  0.80 needs review) and its Evidence - the line as printed. Each statement
  gets a line on the **Review** sheet: opening, credits, debits, net
  movement, computed and printed closing balance, printed totals (and fees,
  where totalled apart), rows needing review, issues, and its final status -
  APPROVED only when every row is and everything reconciles.
- **Nobody is turned away.** Every readable statement is added; anything
  unconfirmed is highlighted for a person to check, not dropped.

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
so glance at the "uncategorised" and "suggested" lists on the upload page
and at the VAT column. VAT is **calculated** at 15% (as instructed), never read from the
statements, and labelled as calculated.

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

**South African banks.** Every bank's statement is read the same way, by
its own running balance. Real statements from ABSA, FNB, Standard Bank and
Capitec are read exactly (tested). Nedbank, Investec, TymeBank, Discovery
Bank, African Bank, Bank Zero and Old Mutual are named in
`config/banks.yaml` and tested on made-up statements; no real statement
from them has been tried yet, so check the first one's Review sheet.

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
  learn.py                   # suggests categories learnt from categorised transactions
  excel_writer.py            # workbook: transactions + report sheets
  gmail_client.py            # OAuth login, search, attachment download
  extract/
    bank_detect.py            # bank + client + statement-period detection
    column_map.py              # header row -> logical column mapping
    text_extract.py             # pdfplumber-based table extraction (primary)
    line_extract.py              # line-by-line fallback when table columns can't be split
    ocr_extract.py                # OCR for scanned PDFs (upright pages, clean-up)
    document.py                    # splits a PDF into statements; verified OCR repairs
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
