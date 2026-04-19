"""Local audit log for outbound email operations (draft creation, etc.)."""

from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path
from typing import Any

from .logging_config import get_logger

logger = get_logger(__name__)

_AUDIT_DIR = Path(os.path.expanduser("~")) / ".outlook-mcp"
_AUDIT_FILE = _AUDIT_DIR / "audit.log"


def log_event(event: str, **fields: Any) -> None:
    try:
        _AUDIT_DIR.mkdir(parents=True, exist_ok=True)
        record = {"ts": datetime.now().isoformat(timespec="seconds"), "event": event, **fields}
        with _AUDIT_FILE.open("a", encoding="utf-8") as f:
            f.write(json.dumps(record, ensure_ascii=False) + "\n")
    except Exception as e:
        logger.warning(f"Audit log write failed: {e}")
