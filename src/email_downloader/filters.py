from datetime import datetime

from email_downloader.models import AttachmentInfo, MailMessage, MailQuery, normalize_extensions


def format_outlook_datetime(value: datetime) -> str:
    """Format a datetime for Outlook Items.Restrict."""
    return value.strftime("%m/%d/%Y %I:%M %p")


def build_restrict_filter(query: MailQuery) -> str | None:
    """Build an Outlook Items.Restrict filter from a mail query."""
    clauses: list[str] = []

    if query.received_after is not None:
        value = format_outlook_datetime(query.received_after)
        clauses.append(f"[ReceivedTime] >= '{value}'")

    if query.received_before is not None:
        value = format_outlook_datetime(query.received_before)
        clauses.append(f"[ReceivedTime] <= '{value}'")

    if query.has_attachment is not None:
        clauses.append(f"[HasAttachment] = {query.has_attachment}")

    if query.unread_only:
        clauses.append("[UnRead] = True")

    return " AND ".join(clauses) or None


def attachment_matches(
    attachment: AttachmentInfo,
    *,
    filename: str | None = None,
    filename_contains: str | None = None,
    extensions: tuple[str, ...] = (),
) -> bool:
    """Return whether one attachment matches all requested conditions."""
    attachment_filename = attachment.filename.casefold()
    # filename 不符合
    if filename is not None and attachment_filename != filename.casefold():
        return False

    # filename_contains 不符合
    if filename_contains is not None and filename_contains.casefold() not in attachment_filename:
        return False

    # 沒有限制 extension
    if not extensions:
        return True

    # 有限制 extension, 檢查是否在允許的副檔名中
    return attachment.extension.casefold() in normalize_extensions(extensions)


def message_matches(message: MailMessage, query: MailQuery) -> bool:
    """Return whether a mail message matches Python-side query conditions."""
    if query.sender is not None:
        if message.sender_email is None:
            return False

        if message.sender_email.casefold() != query.sender.casefold():
            return False

    if query.sender_contains is not None:
        if message.sender_email is None:
            return False

        if query.sender_contains.casefold() not in message.sender_email.casefold():
            return False

    if query.subject is not None and message.subject.casefold() != query.subject.casefold():
        return False

    if (
        query.subject_contains is not None
        and query.subject_contains.casefold() not in message.subject.casefold()
    ):
        return False

    if query.has_attachment is True and not message.attachments:
        return False

    if query.has_attachment is False and message.attachments:
        return False

    has_attachment_filter = (
        query.attachment_name is not None
        or query.attachment_name_contains is not None
        or bool(query.attachment_extensions)
    )

    if has_attachment_filter:
        return any(
            attachment_matches(
                attachment,
                filename=query.attachment_name,
                filename_contains=query.attachment_name_contains,
                extensions=query.attachment_extensions,
            )
            for attachment in message.attachments
        )

    return True
