"""Command-line entry point.

Two commands, both funnelling into the same parser/writer:
  - process-folder: local folder of manually-saved PDFs -> combined workbook.
  - fetch-email: search Gmail for statement-looking PDFs, download the new
    ones, then run them through the exact same pipeline as process-folder.
"""
from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass, field
from pathlib import Path

from . import config as config_mod
from .excel_writer import append_transactions
from .extract.parser import parse_statement
from .models import StatementResult
from .store import ProcessedStore, sha256_of_bytes, sha256_of_file


@dataclass
class RunStats:
    scanned: int = 0
    processed: int = 0
    skipped_already_done: int = 0
    transactions_added: int = 0
    problems: list[StatementResult] = field(default_factory=list)
    reviews: list[StatementResult] = field(default_factory=list)


def _print_summary(stats: RunStats, *, scanned_label: str, workbook: Path) -> None:
    print("\n" + "=" * 60)
    print("RUN SUMMARY")
    print("=" * 60)
    print(f"{scanned_label}:{' ' * max(1, 29 - len(scanned_label))}{stats.scanned}")
    print(f"Already processed (skipped): {stats.skipped_already_done}")
    print(f"New statements processed:    {stats.processed}")
    print(f"Transactions added:          {stats.transactions_added}")

    if stats.reviews:
        print(f"\nFlagged for manual review ({len(stats.reviews)}):")
        for r in stats.reviews:
            reason = r.warning or ("OCR extraction used" if r.used_ocr else "")
            print(f"  - {r.source_file}: {reason}")

    if stats.problems:
        print(f"\nFAILED to parse ({len(stats.problems)}):")
        for r in stats.problems:
            print(f"  - {r.source_file}: {r.error}")

    if not stats.reviews and not stats.problems:
        print("\nNothing needs manual review.")
    print("=" * 60)
    print(f"\nCombined workbook: {workbook}")


def _process_one_pdf(
    pdf_path: Path,
    *,
    source_id: str,
    content_hash: str,
    layouts, generic, client_rules, settings,
    store: ProcessedStore,
    stats: RunStats,
    interactive: bool,
    sender: str | None = None,
    subject: str | None = None,
) -> None:
    print(f"Processing {pdf_path.name} ...")
    result = parse_statement(
        pdf_path,
        layouts=layouts,
        generic=generic,
        client_rules=client_rules,
        settings=settings,
        sender=sender,
        subject=subject,
        interactive=interactive,
    )

    if not result.ok:
        stats.problems.append(result)
        return

    write_result = append_transactions(settings.output_workbook, result.transactions)
    stats.transactions_added += write_result.added
    stats.processed += 1
    store.mark_processed(
        source_id, content_hash, pdf_path.name, result.client, result.bank_display_name,
        len(result.transactions),
    )
    if result.used_ocr or result.warning:
        stats.reviews.append(result)


def cmd_process_folder(args: argparse.Namespace) -> int:
    settings = config_mod.load_settings()
    client_rules = config_mod.load_client_rules(settings.clients_config)
    layouts, generic = config_mod.load_bank_layouts(settings.banks_config)

    folder = Path(args.folder) if args.folder else settings.incoming_pdfs_dir
    if not folder.exists():
        print(f"Folder not found: {folder}", file=sys.stderr)
        return 1

    pdf_paths = sorted(folder.glob("*.pdf"))
    stats = RunStats(scanned=len(pdf_paths))
    if stats.scanned == 0:
        print(f"No PDF files found in {folder}")
        return 0

    with ProcessedStore(settings.processed_db) as store:
        for pdf_path in pdf_paths:
            source_id = f"file:{pdf_path.resolve()}"
            content_hash = sha256_of_file(pdf_path)
            if store.is_processed(source_id, content_hash) and not args.reprocess:
                stats.skipped_already_done += 1
                continue
            _process_one_pdf(
                pdf_path,
                source_id=source_id,
                content_hash=content_hash,
                layouts=layouts, generic=generic, client_rules=client_rules, settings=settings,
                store=store, stats=stats, interactive=not args.no_interactive,
            )

    _print_summary(stats, scanned_label="PDFs scanned", workbook=settings.output_workbook)
    return 0


def cmd_fetch_email(args: argparse.Namespace) -> int:
    from . import gmail_client  # deferred: only needed for this command

    settings = config_mod.load_settings()
    client_rules = config_mod.load_client_rules(settings.clients_config)
    layouts, generic = config_mod.load_bank_layouts(settings.banks_config)
    lookback_days = args.days or settings.gmail_lookback_days

    print("Connecting to Gmail ...")
    service = gmail_client.get_service(settings)
    query = gmail_client.build_search_query(client_rules, lookback_days)
    print(f"Searching Gmail: {query}")
    message_ids = gmail_client.search_message_ids(service, query)

    stats = RunStats()
    settings.email_downloads_dir.mkdir(parents=True, exist_ok=True)

    with ProcessedStore(settings.processed_db) as store:
        for message_id in message_ids:
            info = gmail_client.get_message_info(service, message_id)
            for attachment in info.pdf_attachments:
                stats.scanned += 1
                source_id = f"email:{message_id}:{attachment.filename}"
                if store.has_source_id(source_id) and not args.reprocess:
                    stats.skipped_already_done += 1
                    continue

                data = gmail_client.download_attachment(service, message_id, attachment.attachment_id)
                content_hash = sha256_of_bytes(data)
                if store.is_processed(source_id, content_hash) and not args.reprocess:
                    stats.skipped_already_done += 1
                    continue

                local_name = gmail_client.safe_local_filename(message_id, attachment.filename)
                local_path = settings.email_downloads_dir / local_name
                local_path.write_bytes(data)

                _process_one_pdf(
                    local_path,
                    source_id=source_id,
                    content_hash=content_hash,
                    layouts=layouts, generic=generic, client_rules=client_rules, settings=settings,
                    store=store, stats=stats, interactive=not args.no_interactive,
                    sender=info.sender, subject=info.subject,
                )

    _print_summary(stats, scanned_label="Attachments scanned", workbook=settings.output_workbook)
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

    p_email = sub.add_parser("fetch-email", help="Search Gmail for statement PDFs and process the new ones")
    p_email.add_argument("--days", type=int, help="Lookback window in days (default: GMAIL_LOOKBACK_DAYS or 30)")
    p_email.add_argument("--reprocess", action="store_true", help="Reprocess attachments even if already recorded")
    p_email.add_argument(
        "--no-interactive", action="store_true",
        help="Never prompt for client mapping; unmapped statements are labeled UNMAPPED_CLIENT",
    )
    p_email.set_defaults(func=cmd_fetch_email)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
