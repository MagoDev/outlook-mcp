"""SQLite index for exported emails: metadata + FTS5 full-text search."""

from __future__ import annotations

import sqlite3
from datetime import datetime
from pathlib import Path

_SCHEMA = """
CREATE TABLE IF NOT EXISTS emails (
    entry_id          TEXT PRIMARY KEY,
    folder_path       TEXT NOT NULL,
    message_class     TEXT,
    subject           TEXT,
    sender_name       TEXT,
    sender_email      TEXT,
    to_recipients     TEXT,
    cc_recipients     TEXT,
    received_at       TEXT,
    size              INTEGER,
    attachments_count INTEGER,
    msg_file          TEXT NOT NULL,
    attachments_dir   TEXT,
    exported_at       TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_emails_folder   ON emails(folder_path);
CREATE INDEX IF NOT EXISTS idx_emails_received ON emails(received_at);
CREATE INDEX IF NOT EXISTS idx_emails_sender   ON emails(sender_email);

CREATE VIRTUAL TABLE IF NOT EXISTS emails_fts USING fts5(
    entry_id UNINDEXED,
    subject,
    sender_name,
    sender_email,
    body,
    tokenize='unicode61'
);

CREATE TABLE IF NOT EXISTS extract_runs (
    run_id          INTEGER PRIMARY KEY AUTOINCREMENT,
    started_at      TEXT NOT NULL,
    finished_at     TEXT,
    folders_scanned INTEGER DEFAULT 0,
    emails_exported INTEGER DEFAULT 0,
    emails_skipped  INTEGER DEFAULT 0,
    errors          INTEGER DEFAULT 0,
    status          TEXT
);

CREATE TABLE IF NOT EXISTS export_errors (
    run_id     INTEGER,
    folder     TEXT,
    entry_id   TEXT,
    subject    TEXT,
    error      TEXT,
    occurred_at TEXT,
    FOREIGN KEY (run_id) REFERENCES extract_runs(run_id)
);
"""


def connect(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    conn.executescript(_SCHEMA)
    return conn


def seen_entry_ids(conn: sqlite3.Connection) -> set[str]:
    return {row[0] for row in conn.execute("SELECT entry_id FROM emails")}


def start_run(conn: sqlite3.Connection) -> int:
    cur = conn.execute(
        "INSERT INTO extract_runs (started_at, status) VALUES (?, 'running')",
        (datetime.now().isoformat(timespec="seconds"),),
    )
    conn.commit()
    return cur.lastrowid  # type: ignore[return-value]


def finish_run(conn, run_id, folders_scanned, exported, skipped, errors, status):
    conn.execute(
        """UPDATE extract_runs
           SET finished_at = ?, folders_scanned = ?, emails_exported = ?,
               emails_skipped = ?, errors = ?, status = ?
           WHERE run_id = ?""",
        (
            datetime.now().isoformat(timespec="seconds"),
            folders_scanned,
            exported,
            skipped,
            errors,
            status,
            run_id,
        ),
    )
    conn.commit()


def insert_email(conn: sqlite3.Connection, meta: dict, body: str) -> None:
    conn.execute(
        """INSERT OR IGNORE INTO emails
           (entry_id, folder_path, message_class, subject, sender_name, sender_email,
            to_recipients, cc_recipients, received_at, size, attachments_count,
            msg_file, attachments_dir, exported_at)
           VALUES (:entry_id, :folder_path, :message_class, :subject, :sender_name, :sender_email,
                   :to_recipients, :cc_recipients, :received_at, :size, :attachments_count,
                   :msg_file, :attachments_dir, :exported_at)""",
        meta,
    )
    conn.execute(
        """INSERT INTO emails_fts (entry_id, subject, sender_name, sender_email, body)
           VALUES (?, ?, ?, ?, ?)""",
        (
            meta["entry_id"],
            meta.get("subject") or "",
            meta.get("sender_name") or "",
            meta.get("sender_email") or "",
            (body or "")[:20000],
        ),
    )


def log_error(conn, run_id, folder, entry_id, subject, error) -> None:
    conn.execute(
        """INSERT INTO export_errors (run_id, folder, entry_id, subject, error, occurred_at)
           VALUES (?, ?, ?, ?, ?, ?)""",
        (run_id, folder, entry_id, subject, error, datetime.now().isoformat(timespec="seconds")),
    )
