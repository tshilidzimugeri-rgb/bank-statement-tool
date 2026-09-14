"""Command-line entry point.

Phase 1 (this file, for now): process a folder of manually-saved PDFs into
the combined workbook. Email fetching is wired in on top of this once the
extraction pipeline has been verified against real statements.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import config as config_mod
from .excel_writer import append_transactions
from .extract.parser import parse_statement
from .models import StatementResult
from .store import ProcessedStore, sha256_of_file


def _print_summary(
    *,
    scanned: int,
    processed: int,
    skipped_already_done: int,
    transactions_added: int,
    problems: list[StatementResult],
    reviews: list[StatementResult],
) -> None:
    print("\n" + "=" * 60)
    print("RUN SUMMARY")
    print("=" * 60)
    print(f"PDFs scanned:            {scanned}")
    print(f"Already processed (skipped): {skipped_already_done}")
    print(f"New statements processed:   {processed}")
    print(f"Transactions added:         {transactions_added}")

    if reviews:
        print(f"\nFlagged for manual review ({len(reviews)}):")
        for r in reviews:
            reason = r.warning or ("OCR extraction used" if r.used_ocr else "")
            print(f"  - {r.source_file}: {reason}")

    if problems:
        print(f"\nFAILED to parse ({len(problems)}):")
        for r in problems:
            print(f"  - {r.source_file}: {r.error}")

    if not reviews and not problems:
        print("\nNothing needs manual review.")
    print("=" * 60)


def cmd_process_folder(args: argparse.Namespace) -> int:
    settings = config_mod.load_settings()
    client_rules = config_mod.load_client_rules(settings.clients_config)
    layouts, generic = config_mod.load_bank_layouts(settings.banks_config)

    folder = Path(args.folder) if args.folder else settings.incoming_pdfs_dir
    if not folder.exists():
        print(f"Folder not found: {folder}", file=sys.stderr)
        return 1

    pdf_paths = sorted(folder.glob("*.pdf"))
    scanned = len(pdf_paths)
    if scanned == 0:
        print(f"No PDF files found in {folder}")
        return 0

    store = ProcessedStore(settings.processed_db)
    processed_count = 0
    skipped_already_done = 0
    transactions_added_total = 0
    problems: list[StatementResult] = []
    reviews: list[StatementResult] = []

    try:
        for pdf_path in pdf_paths:
            source_id = f"file:{pdf_path.resolve()}"
            content_hash = sha256_of_file(pdf_path)
            if store.is_processed(source_id, content_hash) and not args.reprocess:
                skipped_already_done += 1
                continue

            print(f"Processing {pdf_path.name} ...")
            result = parse_statement(
                pdf_path,
                layouts=layouts,
                generic=generic,
                client_rules=client_rules,
                settings=settings,
                interactive=not args.no_interactive,
            )

            if not result.ok:
                problems.append(result)
                continue

            write_result = append_transactions(settings.output_workbook, result.transactions)
            transactions_added_total += write_result.added
            processed_count += 1
            store.mark_processed(
                source_id, content_hash, pdf_path.name, result.client, result.bank_display_name,
                len(result.transactions),
            )

            if result.used_ocr or result.warning:
                reviews.append(result)
    finally:
        store.close()

    _print_summary(
        scanned=scanned,
        processed=processed_count,
        skipped_already_done=skipped_already_done,
        transactions_added=transactions_added_total,
        problems=problems,
        reviews=reviews,
    )
    print(f"\nCombined workbook: {settings.output_workbook}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="statement-tool", description="Bank statement PDF -> Excel tool")
    sub = parser.add_subparsers(dest="command", required=True)

    p_folder = sub.add_parser("process-folder", help="Process PDFs from a local folder into the combined workbook")
    p_folder.add_argument("--folder", help="Folder of PDFs (default: data/incoming_pdfs)")
    p_folder.add_argument("--reprocess", action="store_true", help="Reprocess files even if already recorded")
    p_folder.add_argument(
        "--no-interactive", action="store_true",
        help="Never prompt for client mapping; unmapped statements are labeled UNMAPPED_CLIENT",
    )
    p_folder.set_defaults(func=cmd_process_folder)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
