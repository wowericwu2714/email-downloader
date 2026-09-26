from datetime import datetime

from email_downloader.filters import (
    attachment_matches,
    build_restrict_filter,
    message_matches,
)
from email_downloader.models import AttachmentInfo, MailMessage, MailQuery


def test_builds_received_and_boolean_filter() -> None:
    query = MailQuery(
        received_after=datetime(2026, 9, 20, 8, 30),
        received_before=datetime(2026, 9, 21, 18, 0),
        has_attachment=True,
        unread_only=True,
    )

    result = build_restrict_filter(query)

    assert result is not None
    assert "[ReceivedTime] >= '09/20/2026 08:30 AM'" in result
    assert "[ReceivedTime] <= '09/21/2026 06:00 PM'" in result
    assert "[HasAttachment] = True" in result
    assert "[UnRead] = True" in result


def make_message(
    *,
    sender_email: str | None = "warehouse@example.com",
    subject: str = "Daily Inventory Report",
    attachments: tuple[AttachmentInfo, ...] = (
        AttachmentInfo(index=1, filename="Daily Inventory.XLSX", extension=".xlsx"),
    ),
) -> MailMessage:
    return MailMessage(
        entry_id="entry-1",  # Assuming an entry_id is required for MailMessage
        store_id="store-1",  # Assuming a store_id is required for MailMessage
        subject=subject,
        sender_name="Warehouse",
        sender_email=sender_email,
        received_time=datetime(2026, 9, 20, 8, 30),
        unread=False,
        attachments=attachments,
    )


def test_matches_sender_case_insensitively() -> None:
    message = make_message()
    query = MailQuery(sender="WAREHOUSE@EXAMPLE.COM")

    assert message_matches(message, query) is True


def test_matches_sender_contains_case_insensitively() -> None:
    message = make_message()
    query = MailQuery(sender_contains="EXAMPLE")

    assert message_matches(message, query) is True


def test_matches_subject_contains_case_insensitively() -> None:
    message = make_message()
    query = MailQuery(subject_contains="inventory")

    assert message_matches(message, query) is True


def test_attachment_conditions_must_match_same_attachment() -> None:
    message = make_message(
        attachments=(
            AttachmentInfo(index=1, filename="每日庫存.pdf", extension=".pdf"),
            AttachmentInfo(index=2, filename="其他資料.xlsx", extension=".xlsx"),
        )
    )
    query = MailQuery(
        attachment_name_contains="每日庫存",
        attachment_extensions=(".xlsx",),
    )

    assert message_matches(message, query) is False


def test_matches_attachment_when_same_attachment_satisfies_all_conditions() -> None:
    message = make_message()

    query = MailQuery(
        attachment_name_contains="inventory",
        attachment_extensions=(".xlsx",),
    )

    assert message_matches(message, query) is True


def test_has_attachment_false_requires_no_attachments() -> None:
    message = make_message(attachments=())

    query = MailQuery(has_attachment=False)

    assert message_matches(message, query) is True


def test_has_attachment_false_rejects_message_with_attachments() -> None:
    message = make_message()

    query = MailQuery(has_attachment=False)

    assert message_matches(message, query) is False


def test_attachment_matches_combined_conditions() -> None:
    attachment = AttachmentInfo(
        index=1,
        filename="Daily Inventory.XLSX",
        extension=".xlsx",
    )

    assert (
        attachment_matches(
            attachment,
            filename_contains="inventory",
            extensions=(".xlsx",),
        )
        is True
    )


def test_returns_none_when_no_restrict_conditions() -> None:
    query = MailQuery()
    result = build_restrict_filter(query)

    assert result is None


def test_sender_filter_rejects_missing_sender_email() -> None:
    message = make_message(sender_email=None)
    query = MailQuery(sender_contains="example")

    assert message_matches(message, query) is False


def test_exact_subject_must_match_case_insensitively() -> None:
    message = make_message(subject="Daily Inventory Report")
    query = MailQuery(subject="DAILY INVENTORY REPORT")

    assert message_matches(message, query) is True


def test_attachment_matches_rejects_exact_filename_mismatch() -> None:
    attachment = AttachmentInfo(index=1, filename="Daily Inventory.XLSX", extension=".xlsx")

    assert attachment_matches(attachment, filename="other.xlsx") is False


def test_attachment_matches_returns_true_when_no_filters_given() -> None:
    attachment = AttachmentInfo(index=1, filename="Daily Inventory.XLSX", extension=".xlsx")

    assert attachment_matches(attachment) is True


def test_exact_sender_rejects_missing_sender_email() -> None:
    message = make_message(sender_email=None)
    query = MailQuery(sender="warehouse@example.com")

    assert message_matches(message, query) is False


def test_exact_sender_rejects_mismatch() -> None:
    message = make_message()
    query = MailQuery(sender="someone-else@example.com")

    assert message_matches(message, query) is False


def test_sender_contains_rejects_mismatch() -> None:
    message = make_message()
    query = MailQuery(sender_contains="nomatch")

    assert message_matches(message, query) is False


def test_exact_subject_rejects_mismatch() -> None:
    message = make_message(subject="Daily Inventory Report")
    query = MailQuery(subject="Something Else")

    assert message_matches(message, query) is False


def test_subject_contains_rejects_mismatch() -> None:
    message = make_message()
    query = MailQuery(subject_contains="nomatch")

    assert message_matches(message, query) is False


def test_has_attachment_true_rejects_message_without_attachments() -> None:
    message = make_message(attachments=())
    query = MailQuery(has_attachment=True)

    assert message_matches(message, query) is False
