import logging
import os
from pathlib import Path
from typing import Any

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
from email_downloader.folders import iter_folder_tree, resolve_folder
from email_downloader.models import ConflictPolicy, MailMessage, MailQuery
from email_downloader.protocols import ComSessionFactory

logger = logging.getLogger(__name__)


def _search_folder(
    folder: Any,
    query: MailQuery,
    *,
    limit: int | None,
) -> list[MailMessage]:
    """Search a single Outlook folder's Items collection (no subfolder recursion)."""
    results: list[MailMessage] = []

    items = folder.Items
    restrict_filter = build_restrict_filter(query)

    if restrict_filter is not None:
        logger.debug("Applying restrict filter: %s", restrict_filter)
        items = items.Restrict(restrict_filter)

    items.Sort("[ReceivedTime]", True)

    for item in items:
        if getattr(item, "Class", None) != 43:  # 43 corresponds to MailItem
            continue

        try:
            message = mail_item_to_message(item, folder.StoreID)
        except Exception:  # noqa: BLE001, S112 -- skip unreadable item; logging deferred to a later task.
            continue

        if not message_matches(message, query):
            continue

        results.append(message)

        if limit is not None and len(results) >= limit:
            break

    return results


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
        """Search Outlook messages matching the supplied query.

        Args:
            query: Search criteria.
            limit: Maximum number of matching messages to return.

        Returns:
            Matching messages ordered from newest to oldest.

        Raises:
            ValueError: If limit is less than 1.
        """

        if limit is not None and limit < 1:
            raise ValueError("limit must be at least 1")

        logger.debug(
            "Searching Outlook mail: folder=%r, limit=%r, recursive=%r",
            query.folder,
            limit,
            query.recursive,
        )

        with self._session_factory.session() as namespace:
            root_folder = resolve_folder(namespace, query.folder)

            if not query.recursive:
                results = _search_folder(root_folder, query, limit=limit)
            else:
                results = []
                for folder in iter_folder_tree(root_folder):
                    try:
                        results.extend(_search_folder(folder, query, limit=None))
                    except Exception:  # noqa: BLE001 -- skip unreadable subfolder; logging deferred to a later task.
                        logger.warning(
                            "Skipping unreadable Outlook subfolder while searching recursively"
                        )
                        continue
                results.sort(key=lambda message: message.received_time, reverse=True)

                if limit is not None:
                    results = results[:limit]

        logger.debug(
            "Outlook search matched %d message(s)",
            len(results),
        )

        return results

    def find_latest(
        self,
        query: MailQuery,
    ) -> MailMessage | None:
        """Return the newest message matching the query.

        Args:
            query: Search criteria.

        Returns:
            The newest matching message, or None when no message matches.
        """

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
        """Download matching attachments from a previously found message.

        The Outlook item is opened again using its EntryID and StoreID.
        No COM object is stored inside MailMessage.

        Args:
            message: Message returned by search() or find_latest().
            output_dir: Directory where attachments will be saved.
            extensions: Allowed attachment extensions.
            filename: Exact attachment filename filter.
            filename_contains: Partial attachment filename filter.
            conflict: Behaviour when the destination already exists.

        Returns:
            Paths successfully saved.

        Raises:
            AttachmentNotFoundError:
                No attachment metadata matches the filters.
            MailAccessError:
                The original Outlook message or attachment can no longer
                be accessed safely.
            AttachmentSaveError:
                Outlook failed to save an attachment.
        """

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
                raise MailAccessError(f"Failed to access mail item '{message.subject}'") from exc

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
                    logger.info(
                        "Skipping attachment %r for message %r (conflict=%r)",
                        current_filename,
                        message.subject,
                        conflict,
                    )
                    continue

                try:
                    attachment.SaveAsFile(str(destination.resolve()))
                except Exception:  # noqa: BLE001 -- best-effort save; failure is recorded and reported after the loop.
                    failures.append(current_filename)
                    continue

                logger.info("Saved attachment %r to %s", current_filename, destination)

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
        """Find the newest matching message and download its attachments.

        Args:
            query: Mail and attachment search criteria.
            output_dir: Directory where attachments will be saved.
            conflict: Behaviour when a destination file already exists.

        Returns:
            Paths successfully saved.

        Raises:
            MailNotFoundError:
                No message matched the query.
            AttachmentNotFoundError:
                A message matched, but no attachment matched.
        """
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
