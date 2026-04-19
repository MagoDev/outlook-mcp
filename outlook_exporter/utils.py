"""Filesystem-safety helpers for the exporter."""

from __future__ import annotations

import re

_INVALID_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_MULTI_WS = re.compile(r"[\s_-]+")
_RESERVED = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


def slugify(text: str, max_length: int = 50) -> str:
    """Aggressive slug for filename stems (lowercase, dashes)."""
    if not text:
        return "untitled"
    s = _INVALID_CHARS.sub("", text).strip()
    s = _MULTI_WS.sub("-", s).strip("-.")
    s = s.lower()
    return s[:max_length] or "untitled"


def safe_foldername(text: str) -> str:
    """Preserve case+spaces for folder mirror, just strip illegal chars."""
    if not text:
        return "Unknown"
    s = _INVALID_CHARS.sub("", text).strip().rstrip(". ")
    if s.upper() in _RESERVED:
        s = f"_{s}"
    return s or "Unknown"


def sanitize_filename(name: str) -> str:
    """Sanitize an attachment filename, keep extension."""
    if not name:
        return "untitled"
    name = _INVALID_CHARS.sub("_", name).strip().rstrip(". ")
    stem_upper = name.split(".", 1)[0].upper()
    if stem_upper in _RESERVED:
        name = f"_{name}"
    return name or "untitled"
