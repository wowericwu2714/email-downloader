import logging
from collections.abc import Generator
from contextlib import contextmanager
from pathlib import PureWindowsPath
from typing import Any

from email_downloader.exceptions import (
    OutlookConnectionError,
    OutlookUnavailableError,
)
from email_downloader.models import AttachmentInfo, MailMessage

logger = logging.getLogger(__name__)


class Pywin32SessionFactory:
    """Create Outlook MAPI sessions through pywin32."""

    @contextmanager
    def session(self) -> Generator[Any, None, None]:
        """Open an Outlook COM session for the current thread.

        pywin32 is imported lazily so importing this package does not
        require Windows or Outlook until a COM operation is actually used.

        Yields:
            Outlook MAPI namespace.

        Raises:
            OutlookUnavailableError:
                pywin32 is unavailable.
            OutlookConnectionError:
                Outlook could not be opened.
        """
        try:
            import pythoncom
            import win32com.client

        except ImportError as exc:
            raise OutlookUnavailableError(
                "pywin32 and Outlook Desktop are required on Windows"
            ) from exc

        pythoncom.CoInitialize()

        try:
            try:
                application = win32com.client.Dispatch("Outlook.Application")
                namespace = application.GetNamespace("MAPI")
            except Exception as exc:
                raise OutlookConnectionError("Could not open Outlook MAPI session") from exc
            yield namespace
        finally:
            pythoncom.CoUninitialize()


def get_sender_email(item: Any) -> str | None:
    """Return the best available SMTP sender address.

    Exchange internal messages may expose an EX address instead of SMTP.
    In that case, GetExchangeUser().PrimarySmtpAddress is preferred.

    Args:
        item: Outlook MailItem-like object.

    Returns:
        SMTP address when available, otherwise the original sender address.
    """

    original_address = getattr(item, "SenderEmailAddress", None)
    sender_type = getattr(item, "SenderEmailType", None)

    if sender_type != "EX":
        return original_address

    try:
        sender = getattr(item, "Sender", None)
        if sender is None:
            return original_address

        exchange_user = sender.GetExchangeUser()

        if exchange_user is None:
            return original_address

        primary_smtp = getattr(exchange_user, "PrimarySmtpAddress", None)
        return primary_smtp or original_address

    except Exception:  # noqa: BLE001 -- Exchange COM lookup failure must fall back safely.
        logger.debug(
            "Could not resolve Exchange sender SMTP address; using original sender address"
        )
        return original_address


def _safe_getattr(obj: Any, name: str, default: Any = None) -> Any:
    """Return an optional COM attribute, falling back when access fails."""
    try:
        return getattr(obj, name, default)
    except Exception:  # noqa: BLE001 -- Optional COM properties may fail independently.
        return default


def mail_item_to_message(item: Any, store_id: str) -> MailMessage:
    """Convert an Outlook MailItem into a pure Python MailMessage.

    Args:
        item: Outlook MailItem-like object.
        store_id: Outlook store identifier used to reopen the message.

    Returns:
        Pure Python MailMessage.

    Raises:
        AttributeError:
            A required Outlook property is unavailable.
    """
    if getattr(item, "Class", None) != 43:  # olMail
        raise ValueError("Outlook item is not a MailItem")

    # Required fields: intentionally accessed directly.
    entry_id = item.EntryID
    subject = item.Subject
    received_time = item.ReceivedTime

    attachments: list[AttachmentInfo] = []
    attachment_collection = _safe_getattr(item, "Attachments")

    if attachment_collection is not None:
        count = _safe_getattr(attachment_collection, "Count", 0)

        for index in range(1, count + 1):
            attachment = attachment_collection.Item(index)
            filename = attachment.Filename

            attachments.append(
                AttachmentInfo(
                    index=index,
                    filename=filename,
                    extension=PureWindowsPath(filename).suffix.casefold(),
                    size=_safe_getattr(attachment, "Size"),
                )
            )

    return MailMessage(
        entry_id=entry_id,
        store_id=store_id,
        subject=subject,
        sender_name=_safe_getattr(item, "SenderName", ""),
        sender_email=get_sender_email(item),
        received_time=received_time,
        unread=bool(_safe_getattr(item, "UnRead", False)),
        attachments=tuple(attachments),
    )
