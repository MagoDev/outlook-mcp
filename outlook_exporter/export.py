"""Export a single MailItem to .msg + attachments sidecar."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from .utils import safe_foldername, sanitize_filename, slugify

OL_MSG = 3  # olMSG save format (native Outlook)
OL_BY_VALUE = 1  # attachment Type for real files (not embedded objects)
MAIL_MESSAGE_CLASSES = ("IPM.Note", "IPM.Schedule")

# MAPI property for hidden (inline/signature) attachments
PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"


def _is_inline_attachment(att) -> bool:
    """True if the attachment is hidden (signature image, inline HTML image)."""
    try:
        if bool(att.PropertyAccessor.GetProperty(PR_ATTACHMENT_HIDDEN)):
            return True
    except Exception:
        pass
    # Fallback: if it has a ContentID, it's referenced inline
    try:
        cid = getattr(att, "ContentID", "") or ""
        if cid:
            return True
    except Exception:
        pass
    return False


def _sender_email(item) -> str:
    try:
        t = getattr(item, "SenderEmailType", "") or ""
        if t == "EX":
            try:
                return item.Sender.GetExchangeUser().PrimarySmtpAddress or ""
            except Exception:
                pass
        return getattr(item, "SenderEmailAddress", "") or ""
    except Exception:
        return ""


def _received_dt(item) -> datetime | None:
    try:
        rt = item.ReceivedTime
    except Exception:
        return None
    if rt is None:
        return None
    if isinstance(rt, datetime):
        return rt
    try:
        return datetime.fromisoformat(str(rt))
    except Exception:
        return None


def _is_mail(item) -> bool:
    mc = str(getattr(item, "MessageClass", "") or "")
    return any(mc == p or mc.startswith(p + ".") for p in MAIL_MESSAGE_CLASSES)


def export_item(item, folder_path: str, archive_root: Path) -> dict:
    """Export one item. Returns a dict:

    - {"status": "skipped", "reason": "..."} — not a mail item / already on disk
    - {"status": "ok", "meta": {...}, "body": "..."}
    - {"status": "error", "error": "..."}
    """
    try:
        if not _is_mail(item):
            return {"status": "skipped", "reason": "not-mail"}

        entry_id = item.EntryID
        message_class = str(getattr(item, "MessageClass", "") or "")
        subject = (getattr(item, "Subject", "") or "").strip() or "(no subject)"
        sender_name = (getattr(item, "SenderName", "") or "").strip()
        sender_email = _sender_email(item)
        to_recipients = (getattr(item, "To", "") or "").strip()
        cc_recipients = (getattr(item, "CC", "") or "").strip()
        received = _received_dt(item) or datetime.now()
        received_iso = received.isoformat()
        size = int(getattr(item, "Size", 0) or 0)
        att_count = int(getattr(item.Attachments, "Count", 0) or 0)

        folder_parts = [safe_foldername(p) for p in folder_path.split("/") if p]
        year = received.strftime("%Y")
        month = received.strftime("%m")
        dest_dir = archive_root.joinpath(*folder_parts, year, month)
        dest_dir.mkdir(parents=True, exist_ok=True)

        date_prefix = received.strftime("%Y-%m-%d_%H%M")
        from_slug = slugify(sender_name, 30)
        subj_slug = slugify(subject, 50)
        short_id = (entry_id or "noid")[-12:]
        stem = f"{date_prefix}_{from_slug}_{subj_slug}_{short_id}"
        msg_path = dest_dir / f"{stem}.msg"

        if msg_path.exists():
            return {"status": "skipped", "reason": "file-exists", "entry_id": entry_id}

        item.SaveAs(str(msg_path), OL_MSG)

        attachments_dir_rel: str | None = None
        if att_count > 0:
            real = [
                a for a in item.Attachments
                if int(getattr(a, "Type", 0) or 0) == OL_BY_VALUE and not _is_inline_attachment(a)
            ]
            if real:
                att_dir = dest_dir / f"{stem}.attachments"
                att_dir.mkdir(exist_ok=True)
                for att in real:
                    fname = sanitize_filename(getattr(att, "FileName", "") or "unnamed")
                    dest = att_dir / fname
                    if dest.exists():
                        i = 1
                        while True:
                            cand = att_dir / f"{dest.stem}_{i}{dest.suffix}"
                            if not cand.exists():
                                dest = cand
                                break
                            i += 1
                    try:
                        att.SaveAsFile(str(dest))
                    except Exception:
                        continue
                attachments_dir_rel = str(att_dir.relative_to(archive_root)).replace("\\", "/")

        meta = {
            "entry_id": entry_id,
            "folder_path": folder_path,
            "message_class": message_class,
            "subject": subject,
            "sender_name": sender_name,
            "sender_email": sender_email,
            "to_recipients": to_recipients,
            "cc_recipients": cc_recipients,
            "received_at": received_iso,
            "size": size,
            "attachments_count": att_count,
            "msg_file": str(msg_path.relative_to(archive_root)).replace("\\", "/"),
            "attachments_dir": attachments_dir_rel,
            "exported_at": datetime.now().isoformat(timespec="seconds"),
        }
        body = getattr(item, "Body", "") or ""
        return {"status": "ok", "meta": meta, "body": body}

    except Exception as e:
        return {
            "status": "error",
            "error": f"{type(e).__name__}: {e}",
            "entry_id": getattr(item, "EntryID", "?"),
            "subject": getattr(item, "Subject", "?"),
        }
