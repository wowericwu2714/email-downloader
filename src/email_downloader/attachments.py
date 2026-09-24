from pathlib import Path

from email_downloader.exceptions import AttachmentConflictError
from email_downloader.models import AttachmentInfo, normalize_extensions


def select_attachments(
    attachments: tuple[AttachmentInfo, ...],
    *,
    filename: str | None = None,
    filename_contains: str | None = None,
    extensions: tuple[str, ...] = (),
) -> tuple[AttachmentInfo, ...]:
    """Select attachments matching the requested filename and extension filters."""

    normalized_exts = normalize_extensions(extensions)

    selected: list[AttachmentInfo] = []

    for attachment in attachments:
        attachment_filename = attachment.filename.casefold()

        if filename is not None and attachment_filename != filename.casefold():
            continue

        if filename_contains is not None and filename_contains.casefold() not in attachment_filename:
            continue

        if normalized_exts and attachment.extension.casefold() not in normalized_exts:
            continue

        selected.append(attachment)

    return tuple(selected)

def resolve_destination(
    destination: Path, 
    *,
    conflict: str
) -> Path | None:
    """Resolve the destination path according to the conflict policy."""
    if not destination.exists():
        return destination

    if conflict == "overwrite":
        return destination

    if conflict == "skip":
        return None

    if conflict == "error":
        raise AttachmentConflictError(
            f"Attachment destination already exists: {destination}"
        )

    if conflict == "rename":
        counter = 1

        while True:
            candidate = destination.with_name(
                f"{destination.stem}_{counter}{destination.suffix}"
            )

            if not candidate.exists():
                return candidate

            counter += 1
    raise ValueError(f"Unknown conflict policy: {conflict}")