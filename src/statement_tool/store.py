"""SQLite-backed record of statements already processed.

Keyed by (source_id, content_hash):
  - In folder mode, source_id is "file:<absolute path>".
  - In email mode, source_id is "email:<message_id>:<attachment filename>".
content_hash is the sha256 of the PDF bytes, so an edited/replaced file with
the same name is treated as new, but re-running the tool on unchanged input
never re-adds the same transactions.
"""
from __future__ import annotations

import hashlib
import sqlite3
from contextlib import closing
from pathlib import Path


SCHEMA = """
CREATE TABLE IF NOT EXISTS processed_statements (
    source_id TEXT NOT NULL,
    content_hash TEXT NOT NULL,
    source_file TEXT NOT NULL,
    client TEXT,
    bank TEXT,
    transaction_count INTEGER NOT NULL DEFAULT 0,
    processed_at TEXT NOT NULL DEFAULT (datetime('now')),
    PRIMARY KEY (source_id, content_hash)
);
"""


def sha256_of_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


class ProcessedStore:
    def __init__(self, db_path: Path):
        db_path.parent.mkdir(parents=True, exist_ok=True)
        self._conn = sqlite3.connect(db_path)
        with closing(self._conn.cursor()) as cur:
            cur.execute(SCHEMA)
        self._conn.commit()

    def is_processed(self, source_id: str, content_hash: str) -> bool:
        with closing(self._conn.cursor()) as cur:
            cur.execute(
                "SELECT 1 FROM processed_statements WHERE source_id = ? AND content_hash = ?",
                (source_id, content_hash),
            )
            return cur.fetchone() is not None

    def mark_processed(
        self,
        source_id: str,
        content_hash: str,
        source_file: str,
        client: str | None,
        bank: str | None,
        transaction_count: int,
    ) -> None:
        with closing(self._conn.cursor()) as cur:
            cur.execute(
                """
                INSERT OR REPLACE INTO processed_statements
                    (source_id, content_hash, source_file, client, bank, transaction_count)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (source_id, content_hash, source_file, client, bank, transaction_count),
            )
        self._conn.commit()

    def close(self) -> None:
        self._conn.close()

    def __enter__(self) -> "ProcessedStore":
        return self

    def __exit__(self, *exc) -> None:
        self.close()
