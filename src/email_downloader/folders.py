from collections.abc import Iterator
from typing import Any

from email_downloader.exceptions import FolderNotFoundError

OL_FOLDER_INBOX = 6


def split_folder_path(path: str) -> tuple[str, ...]:
    normalized = path.replace("\\", "/")
    parts = tuple(part.strip() for part in normalized.split("/"))

    if not parts or any(not part for part in parts):
        raise ValueError("folder path contains an empty segment")

    if parts[0].casefold() in {"inbox", "收件匣"}:
        parts = parts[1:]
    return parts


def resolve_folder(namespace: Any, path: str) -> Any:
    parts = split_folder_path(path)
    folder = namespace.GetDefaultFolder(OL_FOLDER_INBOX)

    try:
        for part in parts:
            folder = folder.Folders.Item(part)
    except Exception as exc:
        raise FolderNotFoundError(f"Outlook folder not found: {path}") from exc

    return folder


def iter_folder_tree(folder: Any) -> Iterator[Any]:
    """Yield a folder and all of its subfolders, depth-first, at every level."""
    yield folder
    for sub in folder.Folders:
        yield from iter_folder_tree(sub)
