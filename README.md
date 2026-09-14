# statement-tool

Turns client bank statement PDFs into one combined Excel workbook, and can
pull the PDFs straight from a Gmail inbox on demand. Built incrementally:

- **Phase 1 (done):** folder of PDFs -> combined `.xlsx`.
- **Phase 2 (done):** Gmail search/download layered on top of the same
  parser, using the Gmail API with OAuth.

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

Two YAML files, both meant to be hand-edited - no code changes needed to
onboard a new client or a straightforward new bank:

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

## Running (phase 1: local folder)

Drop PDFs into `data/incoming_pdfs/` (or point `--folder` elsewhere), then:

```powershell
.venv\Scripts\python -m statement_tool.cli process-folder
```

- Already-processed PDFs (tracked by file hash in
  `data/processed/processed.db`) are skipped automatically on the next run.
- New transactions are appended to `output/combined_statements.xlsx` on a
  single `Transactions` sheet, formatted as an Excel Table with a `TOTAL`
  row driven by real `=SUM(...)` formulas (they recalculate if you edit a
  cell), plus a `Summary` sheet with per-client/per-bank `SUMIFS` totals,
  both fully rebuilt each run from whatever is currently in `Transactions`.
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

Copy a block in `config/banks.yaml`, rename the key, and adjust:

- `detect` - a few substrings unique to that bank's statement header/footer.
- `columns` - the header text aliases pdfplumber will see in that bank's
  table headers.
- `amount_style` - `split` (separate debit/credit columns, most SA banks)
  or `signed` (one amount column, negative = debit).

No Python changes needed for a normal text-based statement with a
recognizable header row. If a bank's layout doesn't have a clean header row
pdfplumber can find, that's a case for extending `text_extract.py`/
`column_map.py` - flag it and we'll add a heuristic for that specific shape.

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
  clients.yaml          # client name <- sender/subject/filename/text matching (edit me)
  banks.yaml             # per-bank column layout + detection (edit me)
src/statement_tool/
  cli.py                 # command-line entry point
  config.py               # loads the YAML + .env settings above
  models.py                # Transaction / StatementResult dataclasses
  store.py                  # SQLite "already processed" tracking
  excel_writer.py            # combined workbook read/append/format
  gmail_client.py            # OAuth login, search, attachment download
  extract/
    bank_detect.py            # bank + client + statement-period detection
    column_map.py              # header row -> logical column mapping
    text_extract.py             # pdfplumber-based table extraction (primary)
    ocr_extract.py                # pytesseract fallback for scanned PDFs
    amounts.py, dates.py           # shared value parsing
    parser.py                       # orchestrates the above per PDF
data/incoming_pdfs/       # drop PDFs here for phase 1
data/email_downloads/      # PDFs fetched via Gmail land here
data/processed/            # processed.db (SQLite dedup record)
output/                      # combined_statements.xlsx lands here
tests/
```

## Credentials

Phase 1 needs no credentials. Phase 2 stores the Gmail OAuth client secret
and token cache as local files referenced from `.env` (`credentials.json`,
`token.json` by default) - both are already in `.gitignore`, along with
`.env` itself and the generated workbook/database, so none of it reaches
source control.
