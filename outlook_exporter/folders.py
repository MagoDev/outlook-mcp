"""Walk Outlook folders, respecting exclusions."""

from __future__ import annotations

from typing import Iterable, Iterator, Tuple


def walk(
    path: str,
    folder,
    excluded_names: Iterable[str],
) -> Iterator[Tuple[str, object]]:
    """Recursively yield (path, Folder) pairs. Case-insensitive name exclusion."""
    excl = {n.lower() for n in excluded_names}
    if folder.Name.lower() in excl:
        return
    yield (path, folder)
    try:
        subfolders = list(folder.Folders)
    except Exception:
        return
    for sub in subfolders:
        try:
            yield from walk(f"{path}/{sub.Name}", sub, excluded_names)
        except Exception:
            continue


def is_mail_folder(folder) -> bool:
    """True if the folder is expected to hold IPM.Note items.

    Calendar/Contacts/Tasks etc. have different DefaultMessageClass values.
    """
    try:
        cls = str(getattr(folder, "DefaultMessageClass", "") or "")
    except Exception:
        return True  # err on the side of walking
    return cls in ("", "IPM.Note", "IPM.Post", "IPM.Note.SMIME", "IPM.Note.SMIME.MultipartSigned")
