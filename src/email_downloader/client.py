import os
from pathlib import Path

from email_downloader.attachments import (
    resolve_destination,
    safe_attachment_name,
    select_attachments,
)
from email_downloader.com_backend import (
    Pywin32SessionFactory,
    mail_item_to_message,
)
from email_downloader.exceptions import (
    AttachmentNotFoundError,
    AttachmentSaveError,
    MailAccessError,
    MailNotFoundError,
)
from email_downloader.filters import (
    build_restrict_filter,
    message_matches,
)
from email_downloader.folders import resolve_folder
from email_downloader.models import ConflictPolicy, MailMessage, MailQuery
from email_downloader.protocols import ComSessionFactory


class OutlookClient:
    """High-level Outlook mail search client."""

    def __init__(
        self, 
        session_factory: ComSessionFactory | None = None,
    ) -> None:
        self._session_factory = session_factory or Pywin32SessionFactory()

    def search(
        self, 
        query: MailQuery,
        *, 
        limit: int | None = None,
    ) -> list[MailMessage]:
        """Search Outlook messages matching the query."""
        if limit is not None and limit < 1:
            raise ValueError("limit must be at least 1")

        results: list[MailMessage] = []

        with self._session_factory.session() as namespace:
            folder = resolve_folder(namespace, query.folder)
            items = folder.Items

            restrict_filter = build_restrict_filter(query)

            if restrict_filter is not None:
                items = items.Restrict(restrict_filter)

            items.Sort("[ReceivedTime]", True)

            for item in items:
                if getattr(item, "Class", None) != 43:  # 43 corresponds to MailItem
                    continue

                try:
                    message = mail_item_to_message(
                        item,
                        folder.StoreID,
                    )
                except Exception:  # noqa: BLE001, S112 -- skip unreadable item; logging deferred to a later task.
                    continue

                if not message_matches(message, query):
                    continue

                results.append(message)

                if limit is not None and len(results) >= limit:
                    break

        return results

    def find_latest(
        self,
        query: MailQuery,
    ) -> MailMessage | None:
        """Return the newest message matching the query."""
        messages = self.search(query, limit=1)
        return messages[0] if messages else None

    def download_attachments(
        self, 
        message: MailMessage,
        output_dir: str | os.PathLike[str],
        *, 
        extensions: tuple[str, ...] = (),
        filename: str | None = None,
        filename_contains: str | None = None,
        conflict: ConflictPolicy = "error",
    ) -> list[Path]:
        """Download matching attachments from a previously found Outlook message."""
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)

        selected = select_attachments(
            message.attachments,
            filename=filename,
            filename_contains=filename_contains,
            extensions=extensions,
        )

        if not selected:
            raise AttachmentNotFoundError(
                f"No attachments found matching the specified criteria in message '{message.subject}'"
            )

        paths: list[Path] = []
        failures: list[str] = []

        with self._session_factory.session() as namespace:
            try:
                item = namespace.GetItemFromID(
                    message.entry_id,
                    message.store_id,
                )
            except Exception as exc:
                raise MailAccessError(
                    f"Failed to access mail item '{message.subject}'"
                ) from exc

            attachments = item.Attachments
            for attachment_info in selected:
                try: 
                    attachment = attachments.Item(attachment_info.index)
                except Exception as exc:
                    raise MailAccessError(
                        f"Attachment at index {attachment_info.index} is no longer available "
                        f"on message {message.subject!r}"
                    ) from exc

                current_filename = safe_attachment_name(attachment.FileName)
                expected_filename = safe_attachment_name(attachment_info.filename)
                

                if current_filename.casefold() != expected_filename.casefold():
                    raise MailAccessError(
                        "Outlook attachment changed after search: "
                        f"expected {attachment_info.filename!r}, "
                        f"got {attachment.FileName!r}"
                    )

                destination = resolve_destination(
                    output_path / current_filename,
                    conflict=conflict,
                )

                if destination is None:
                    continue

                try:
                    attachment.SaveAsFile(str(destination.resolve()))
                except Exception: # noqa: BLE001 -- best-effort save; failure is recorded and reported after the loop.
                    failures.append(current_filename)


                paths.append(destination)

            if failures:
                raise AttachmentSaveError(
                    f"Failed to save {len(failures)} attachment(s): {failures}; "
                    f"{len(paths)} attachment(s) already saved: {paths}"
                )

        return paths


    def download_latest(
        self, 
        query: MailQuery,
        output_dir: str | os.PathLike[str],
        *,
        conflict: ConflictPolicy = "error",
    ) -> list[Path]:
        """Download matching attachments from the latest matching message."""
        message = self.find_latest(query)
        
        if message is None:
            raise MailNotFoundError(f"No mail found matching query: {query}")

        return self.download_attachments(
            message,
            output_dir,
            extensions=query.attachment_extensions,
            filename=query.attachment_name,
            filename_contains=query.attachment_name_contains,
            conflict=conflict,
        )