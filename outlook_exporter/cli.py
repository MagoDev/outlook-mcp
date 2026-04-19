"""Command-line interface for the Outlook exporter."""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

import pythoncom
import win32com.client

from . import db as dbmod
from . import export as exportmod
from .folders import is_mail_folder, walk

DEFAULT_ARCHIVE = Path.home() / "SynologyDrive" / "Emails" / "archive"
DEFAULT_DB = Path.home() / "SynologyDrive" / "Emails" / "index" / "emails.sqlite"
DEFAULT_EXCLUDES = ["Deleted Items"]
COMMIT_EVERY = 50


def _parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        prog="outlook-export",
        description="Incrementally export Outlook Classic mail to .msg files + SQLite index.",
    )
    p.add_argument("--archive", type=Path, default=DEFAULT_ARCHIVE,
                   help=f"Archive root (default: {DEFAULT_ARCHIVE})")
    p.add_argument("--db", type=Path, default=DEFAULT_DB,
                   help=f"SQLite index path (default: {DEFAULT_DB})")
    p.add_argument("--mailbox", type=str, default=None,
                   help="Mailbox display name to export (default: first account found)")
    p.add_argument("--folder", type=str, default=None,
                   help='Limit to a sub-path, e.g. "Inbox/EAO". Default: all folders on the mailbox.')
    p.add_argument("--exclude", action="append", default=list(DEFAULT_EXCLUDES),
                   help="Folder name to exclude (repeatable). Default: Deleted Items.")
    p.add_argument("--dry-run", action="store_true",
                   help="List folders+counts without exporting.")
    p.add_argument("--max-items", type=int, default=None,
                   help="Stop after exporting N items (for testing).")
    p.add_argument("--verbose", action="store_true")
    return p.parse_args()


def _get_mailbox(mapi, wanted: str | None):
    stores = list(mapi.Folders)
    if not stores:
        raise SystemExit("No Outlook mailboxes available.")
    if wanted:
        for s in stores:
            if s.Name == wanted:
                return s
        names = ", ".join(s.Name for s in stores)
        raise SystemExit(f"Mailbox '{wanted}' not found. Available: {names}")
    return stores[0]


def _resolve_folder(root, subpath: str | None):
    if not subpath:
        return root
    folder = root
    for part in subpath.split("/"):
        folder = folder.Folders[part]
    return folder


def main() -> int:
    args = _parse_args()
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mapi = outlook.GetNamespace("MAPI")
        mailbox = _get_mailbox(mapi, args.mailbox)
        print(f"Mailbox: {mailbox.Name}")
        root = _resolve_folder(mailbox, args.folder)
        root_path = f"{mailbox.Name}" + (f"/{args.folder}" if args.folder else "")

        conn = dbmod.connect(args.db)
        seen = dbmod.seen_entry_ids(conn)
        print(f"DB: {args.db} ({len(seen)} emails already indexed)")

        folders = list(walk(root_path, root, args.exclude))
        print(f"Folders to scan: {len(folders)}")

        if args.dry_run:
            total = 0
            for path, folder in folders:
                if not is_mail_folder(folder):
                    print(f"  SKIP (non-mail): {path}")
                    continue
                try:
                    n = folder.Items.Count
                except Exception:
                    n = "?"
                print(f"  {path} — {n} items")
                if isinstance(n, int):
                    total += n
            print(f"Total items across mail folders: {total}")
            return 0

        run_id = dbmod.start_run(conn)
        args.archive.mkdir(parents=True, exist_ok=True)

        exported = 0
        skipped = 0
        errors = 0
        start = time.time()

        try:
            for fpath, folder in folders:
                if not is_mail_folder(folder):
                    if args.verbose:
                        print(f"SKIP non-mail: {fpath}")
                    continue
                try:
                    total_in_folder = folder.Items.Count
                except Exception:
                    total_in_folder = 0
                print(f"[{fpath}] {total_in_folder} items")

                # Iterate via positional access to be resilient to collection changes.
                items = folder.Items
                for idx in range(1, total_in_folder + 1):
                    try:
                        item = items.Item(idx)
                    except Exception as e:
                        errors += 1
                        dbmod.log_error(conn, run_id, fpath, "?", "?", f"item-fetch: {e}")
                        continue

                    try:
                        eid = item.EntryID
                    except Exception:
                        eid = None
                    if eid and eid in seen:
                        skipped += 1
                        continue

                    result = exportmod.export_item(item, fpath, args.archive)
                    st = result["status"]
                    if st == "ok":
                        dbmod.insert_email(conn, result["meta"], result["body"])
                        seen.add(result["meta"]["entry_id"])
                        exported += 1
                        if args.verbose:
                            print(f"  + {result['meta']['msg_file']}")
                        if exported % COMMIT_EVERY == 0:
                            conn.commit()
                            rate = exported / max(time.time() - start, 1e-3)
                            print(f"  … {exported} exported ({rate:.1f}/s)")
                        if args.max_items and exported >= args.max_items:
                            print(f"Reached --max-items={args.max_items}, stopping.")
                            raise StopIteration
                    elif st == "skipped":
                        skipped += 1
                        if eid:
                            seen.add(eid)
                    elif st == "error":
                        errors += 1
                        dbmod.log_error(
                            conn, run_id, fpath,
                            result.get("entry_id", "?"),
                            result.get("subject", "?"),
                            result.get("error", "?"),
                        )
                        if args.verbose:
                            print(f"  ! ERROR {result.get('error')}")
        except StopIteration:
            pass
        except KeyboardInterrupt:
            print("\nInterrupted by user — partial progress is saved.")
            dbmod.finish_run(conn, run_id, len(folders), exported, skipped, errors, "interrupted")
            conn.commit()
            return 130

        dbmod.finish_run(conn, run_id, len(folders), exported, skipped, errors, "ok")
        conn.commit()
        elapsed = time.time() - start
        print(
            f"\nDone. exported={exported} skipped={skipped} errors={errors} "
            f"elapsed={elapsed:.1f}s ({exported / max(elapsed, 1e-3):.1f} items/s)"
        )
        print(f"Archive: {args.archive}")
        print(f"Index:   {args.db}")
        return 0 if errors == 0 else 1

    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    sys.exit(main())
