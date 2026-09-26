from collections.abc import Iterator
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Any
from unittest.mock import Mock

import pytest

import email_downloader.client as client_module
from email_downloader.client import OutlookClient
from email_downloader.exceptions import (
    AttachmentNotFoundError,
    AttachmentSaveError,
    MailAccessError,
    MailNotFoundError,
)
from email_downloader.models import AttachmentInfo, MailMessage, MailQuery


class FakeItems:
    def __init__(self, items: list[Any]) -> None:
        self._items = items
        self.restrict_calls: list[str] = []
        self.sort_calls: list[tuple[str, bool]] = []

    def Restrict(self, filter_string: str) -> "FakeItems":
        self.restrict_calls.append(filter_string)
        return self

    def Sort(self, property_name: str, descending: bool) -> None:
        self.sort_calls.append((property_name, descending))

    def __iter__(self) -> Iterator[Any]:
        return iter(self._items)


class FakeFolder:
    def __init__(self, items: FakeItems, folders: list[Any] | None = None) -> None:
        self.Items = items
        self.StoreID = "store-123"
        self.Folders = folders or []


class FakeSessionFactory:
    def __init__(self, namespace: object) -> None:
        self.namespace = namespace

    @contextmanager
    def session(self) -> Iterator[object]:
        yield self.namespace


def make_message(
    entry_id: str,
    *,
    subject: str = "Daily Inventory",
    received_time: datetime = datetime(2026, 9, 23, 8, 0),
) -> MailMessage:
    return MailMessage(
        entry_id=entry_id,
        store_id="store-123",
        subject=subject,
        sender_name="Warehouse",
        sender_email="warehouse@example.com",
        received_time=received_time,
        unread=False,
        attachments=(),
    )


def test_search_restricts_sorts_and_returns_matching_messages(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class FakeMailItem:
        Class = 43

    mail_item = FakeMailItem()

    items = FakeItems([mail_item])
    folder = FakeFolder(items)
    namespace = object()

    message = make_message(entry_id="entry-1")

    monkeypatch.setattr(
        client_module,
        "resolve_folder",
        lambda namespace, path: folder,
    )
    monkeypatch.setattr(
        client_module,
        "build_restrict_filter",
        lambda query: "[HasAttachment] = True",
    )
    monkeypatch.setattr(
        client_module,
        "mail_item_to_message",
        lambda item, store_id: message,
    )
    monkeypatch.setattr(
        client_module,
        "message_matches",
        lambda message, query: True,
    )

    client = OutlookClient(session_factory=FakeSessionFactory(namespace))

    result = client.search(MailQuery(has_attachment=True))

    assert result == [message]

    assert items.restrict_calls == ["[HasAttachment] = True"]

    assert items.sort_calls == [("[ReceivedTime]", True)]


def test_search_does_not_restrict_when_filter_is_none(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    items = FakeItems([])
    folder = FakeFolder(items)
    namespace = object()

    monkeypatch.setattr(
        client_module,
        "resolve_folder",
        lambda namespace, path: folder,
    )
    monkeypatch.setattr(
        client_module,
        "build_restrict_filter",
        lambda query: None,
    )

    client = OutlookClient(session_factory=FakeSessionFactory(namespace))

    result = client.search(MailQuery())

    assert result == []
    assert items.restrict_calls == []

    assert items.sort_calls == [("[ReceivedTime]", True)]


@pytest.mark.parametrize("limit", [0, -1])
def test_search_rejects_invalid_limit(limit: int) -> None:
    client = OutlookClient(session_factory=FakeSessionFactory(object()))

    with pytest.raises(
        ValueError,
        match="limit must be at least 1",
    ):
        client.search(MailQuery(), limit=limit)


def test_search_ignores_non_mail_items(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class FakeNonMailItem:
        Class = 99

    items = FakeItems([FakeNonMailItem()])
    folder = FakeFolder(items)
    namespace = object()

    monkeypatch.setattr(
        client_module,
        "resolve_folder",
        lambda namespace, path: folder,
    )
    monkeypatch.setattr(
        client_module,
        "build_restrict_filter",
        lambda query: None,
    )

    client = OutlookClient(session_factory=FakeSessionFactory(namespace))

    result = client.search(MailQuery())

    assert result == []


def test_search_excludes_messages_rejected_by_python_matcher(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class FakeMailItem:
        Class = 43

    items = FakeItems([FakeMailItem()])
    folder = FakeFolder(items)
    namespace = object()

    message = make_message("entry-1")

    monkeypatch.setattr(
        client_module,
        "resolve_folder",
        lambda namespace, path: folder,
    )
    monkeypatch.setattr(
        client_module,
        "build_restrict_filter",
        lambda query: None,
    )
    monkeypatch.setattr(
        client_module,
        "mail_item_to_message",
        lambda item, store_id: message,
    )
    monkeypatch.setattr(
        client_module,
        "message_matches",
        lambda message, query: False,
    )

    client = OutlookClient(session_factory=FakeSessionFactory(namespace))

    result = client.search(MailQuery())

    assert result == []


def test_search_stops_when_limit_is_reached(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class FakeMailItem:
        Class = 43

    mail_items = [
        FakeMailItem(),
        FakeMailItem(),
        FakeMailItem(),
    ]

    items = FakeItems(mail_items)
    folder = FakeFolder(items)
    namespace = object()

    messages = iter(
        [
            make_message("entry-1"),
            make_message("entry-2"),
            make_message("entry-3"),
        ]
    )

    monkeypatch.setattr(
        client_module,
        "resolve_folder",
        lambda namespace, path: folder,
    )
    monkeypatch.setattr(
        client_module,
        "build_restrict_filter",
        lambda query: None,
    )
    monkeypatch.setattr(
        client_module,
        "mail_item_to_message",
        lambda item, store_id: next(messages),
    )
    monkeypatch.setattr(
        client_module,
        "message_matches",
        lambda message, query: True,
    )

    client = OutlookClient(session_factory=FakeSessionFactory(namespace))

    result = client.search(
        MailQuery(),
        limit=2,
    )

    assert [message.entry_id for message in result] == [
        "entry-1",
        "entry-2",
    ]


def test_search_skips_items_that_fail_to_convert(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class FakeMailItem:
        Class = 43

    broken_item = FakeMailItem()
    ok_item = FakeMailItem()

    items = FakeItems([broken_item, ok_item])
    folder = FakeFolder(items)
    namespace = object()

    message = make_message("entry-1")

    def fake_mail_item_to_message(item: Any, store_id: str) -> MailMessage:
        if item is broken_item:
            raise RuntimeError("item was moved or deleted")
        return message

    monkeypatch.setattr(
        client_module,
        "resolve_folder",
        lambda namespace, path: folder,
    )
    monkeypatch.setattr(
        client_module,
        "build_restrict_filter",
        lambda query: None,
    )
    monkeypatch.setattr(
        client_module,
        "mail_item_to_message",
        fake_mail_item_to_message,
    )
    monkeypatch.setattr(
        client_module,
        "message_matches",
        lambda message, query: True,
    )

    client = OutlookClient(session_factory=FakeSessionFactory(namespace))

    result = client.search(MailQuery())

    assert result == [message]


def test_find_latest_returns_first_search_result(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    client = OutlookClient(session_factory=FakeSessionFactory(object()))

    message = make_message("entry-1")

    monkeypatch.setattr(
        client,
        "search",
        lambda query, *, limit=None: [message],
    )

    result = client.find_latest(MailQuery())

    assert result == message


def test_find_latest_returns_none_when_no_message_matches(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    client = OutlookClient(session_factory=FakeSessionFactory(object()))

    monkeypatch.setattr(
        client,
        "search",
        lambda query, *, limit=None: [],
    )

    result = client.find_latest(MailQuery())

    assert result is None


def test_download_attachments_returns_saved_paths(
    tmp_path: Path,
) -> None:
    fake_attachment = Mock()
    fake_attachment.FileName = "每日庫存.xlsx"

    fake_attachments = Mock()
    fake_attachments.Item.return_value = fake_attachment

    fake_mail_item = Mock()
    fake_mail_item.Attachments = fake_attachments

    namespace = Mock()
    namespace.GetItemFromID.return_value = fake_mail_item

    session_factory = Mock()
    session_factory.session.return_value.__enter__ = Mock(return_value=namespace)
    session_factory.session.return_value.__exit__ = Mock(return_value=False)

    message = MailMessage(
        entry_id="entry-001",
        store_id="store-001",
        subject="每日庫存",
        sender_name="Warehouse",
        sender_email="warehouse@example.com",
        received_time=__import__("datetime").datetime(2026, 9, 24, 8, 0),
        unread=False,
        attachments=(
            AttachmentInfo(
                index=1,
                filename="每日庫存.xlsx",
                extension=".xlsx",
            ),
        ),
    )

    client = OutlookClient(session_factory=session_factory)

    paths = client.download_attachments(
        message,
        output_dir=tmp_path,
        extensions=(".xlsx",),
        conflict="error",
    )

    expected = tmp_path / "每日庫存.xlsx"

    assert paths == [expected]
    namespace.GetItemFromID.assert_called_once_with(
        message.entry_id,
        message.store_id,
    )

    fake_attachments.Item.assert_called_once_with(1)
    fake_attachment.SaveAsFile.assert_called_once_with(str(expected.resolve()))


def test_download_attachments_translates_get_item_failure(
    tmp_path: Path,
) -> None:
    namespace = Mock()
    namespace.GetItemFromID.side_effect = RuntimeError("mail moved")

    session_factory = Mock()
    session_factory.session.return_value.__enter__ = Mock(return_value=namespace)
    session_factory.session.return_value.__exit__ = Mock(return_value=False)

    message = MailMessage(
        entry_id="entry-001",
        store_id="store-001",
        subject="每日庫存",
        sender_name="Warehouse",
        sender_email="warehouse@example.com",
        received_time=datetime(2026, 9, 24, 8, 0),
        unread=False,
        attachments=(AttachmentInfo(1, "每日庫存.xlsx", ".xlsx"),),
    )

    client = OutlookClient(session_factory=session_factory)

    with pytest.raises(MailAccessError):
        client.download_attachments(
            message,
            output_dir=tmp_path,
            extensions=(".xlsx",),
        )


def test_download_attachments_raises_when_no_attachment_matches(
    tmp_path: Path,
) -> None:
    session_factory = Mock()
    client = OutlookClient(session_factory=session_factory)

    message = MailMessage(
        entry_id="entry-001",
        store_id="store-001",
        subject="每日庫存",
        sender_name="Warehouse",
        sender_email="warehouse@example.com",
        received_time=datetime(2026, 9, 24, 8, 0),
        unread=False,
        attachments=(AttachmentInfo(1, "每日庫存.pdf", ".pdf"),),
    )

    with pytest.raises(AttachmentNotFoundError):
        client.download_attachments(
            message,
            output_dir=tmp_path,
            extensions=(".xlsx",),
        )


def test_download_attachments_translates_save_failure(
    tmp_path: Path,
) -> None:
    fake_attachment = Mock()
    fake_attachment.FileName = "每日庫存.xlsx"
    fake_attachment.SaveAsFile.side_effect = RuntimeError("save failed")

    fake_attachments = Mock()
    fake_attachments.Item.return_value = fake_attachment

    fake_mail_item = Mock()
    fake_mail_item.Attachments = fake_attachments

    namespace = Mock()
    namespace.GetItemFromID.return_value = fake_mail_item

    session_factory = Mock()
    session_factory.session.return_value.__enter__ = Mock(return_value=namespace)
    session_factory.session.return_value.__exit__ = Mock(return_value=False)

    message = MailMessage(
        entry_id="entry-001",
        store_id="store-001",
        subject="每日庫存",
        sender_name="Warehouse",
        sender_email="warehouse@example.com",
        received_time=datetime(2026, 9, 24, 8, 0),
        unread=False,
        attachments=(AttachmentInfo(1, "每日庫存.xlsx", ".xlsx"),),
    )

    client = OutlookClient(session_factory=session_factory)

    with pytest.raises(AttachmentSaveError, match="每日庫存.xlsx"):
        client.download_attachments(
            message,
            output_dir=tmp_path,
            extensions=(".xlsx",),
        )


def test_download_attachments_saves_remaining_when_one_fails(
    tmp_path: Path,
) -> None:
    fake_attachment_ok = Mock()
    fake_attachment_ok.FileName = "每日庫存.xlsx"

    fake_attachment_fail = Mock()
    fake_attachment_fail.FileName = "備註.pdf"
    fake_attachment_fail.SaveAsFile.side_effect = RuntimeError("disk full")

    fake_attachments = Mock()
    fake_attachments.Item.side_effect = lambda index: {
        1: fake_attachment_ok,
        2: fake_attachment_fail,
    }[index]

    fake_mail_item = Mock()
    fake_mail_item.Attachments = fake_attachments

    namespace = Mock()
    namespace.GetItemFromID.return_value = fake_mail_item

    session_factory = Mock()
    session_factory.session.return_value.__enter__ = Mock(return_value=namespace)
    session_factory.session.return_value.__exit__ = Mock(return_value=False)

    message = MailMessage(
        entry_id="entry-001",
        store_id="store-001",
        subject="每日庫存",
        sender_name="Warehouse",
        sender_email="warehouse@example.com",
        received_time=datetime(2026, 9, 24, 8, 0),
        unread=False,
        attachments=(
            AttachmentInfo(index=1, filename="每日庫存.xlsx", extension=".xlsx"),
            AttachmentInfo(index=2, filename="備註.pdf", extension=".pdf"),
        ),
    )

    client = OutlookClient(session_factory=session_factory)

    with pytest.raises(AttachmentSaveError, match="備註.pdf") as exc_info:
        client.download_attachments(
            message,
            output_dir=tmp_path,
        )

    assert "每日庫存.xlsx" in str(exc_info.value)

    fake_attachment_ok.SaveAsFile.assert_called_once_with(
        str((tmp_path / "每日庫存.xlsx").resolve())
    )
    fake_attachment_fail.SaveAsFile.assert_called_once()


def test_download_attachments_skip_all_returns_empty_list(
    tmp_path: Path,
) -> None:

    existing = tmp_path / "每日庫存.xlsx"
    existing.touch()

    fake_attachment = Mock()
    fake_attachment.FileName = "每日庫存.xlsx"

    fake_attachments = Mock()
    fake_attachments.Item.return_value = fake_attachment

    fake_mail_item = Mock()
    fake_mail_item.Attachments = fake_attachments

    namespace = Mock()
    namespace.GetItemFromID.return_value = fake_mail_item

    session_factory = Mock()
    session_factory.session.return_value.__enter__ = Mock(return_value=namespace)
    session_factory.session.return_value.__exit__ = Mock(return_value=False)

    message = MailMessage(
        entry_id="entry-001",
        store_id="store-001",
        subject="每日庫存",
        sender_name="Warehouse",
        sender_email="warehouse@example.com",
        received_time=datetime(2026, 9, 24, 8, 0),
        unread=False,
        attachments=(AttachmentInfo(1, "每日庫存.xlsx", ".xlsx"),),
    )

    client = OutlookClient(session_factory=session_factory)

    paths = client.download_attachments(
        message,
        output_dir=tmp_path,
        extensions=(".xlsx",),
        conflict="skip",
    )

    assert paths == []
    fake_attachment.SaveAsFile.assert_not_called()


def test_download_attachments_raises_when_attachment_changed(
    tmp_path: Path,
) -> None:
    fake_attachment = Mock()
    fake_attachment.FileName = "錯誤檔案.xlsx"

    fake_attachments = Mock()
    fake_attachments.Item.return_value = fake_attachment

    fake_mail_item = Mock()
    fake_mail_item.Attachments = fake_attachments

    namespace = Mock()
    namespace.GetItemFromID.return_value = fake_mail_item

    session_factory = Mock()
    session_factory.session.return_value.__enter__ = Mock(return_value=namespace)
    session_factory.session.return_value.__exit__ = Mock(return_value=False)

    message = MailMessage(
        entry_id="entry-001",
        store_id="store-001",
        subject="每日庫存",
        sender_name="Warehouse",
        sender_email="warehouse@example.com",
        received_time=datetime(2026, 9, 24, 8, 0),
        unread=False,
        attachments=(
            AttachmentInfo(
                index=1,
                filename="每日庫存.xlsx",
                extension=".xlsx",
            ),
        ),
    )

    client = OutlookClient(session_factory=session_factory)

    with pytest.raises(MailAccessError, match="每日庫存.xlsx"):
        client.download_attachments(
            message,
            output_dir=tmp_path,
            extensions=(".xlsx",),
        )

    fake_attachment.SaveAsFile.assert_not_called()


def test_download_attachments_raises_when_attachment_missing(
    tmp_path: Path,
) -> None:
    fake_attachments = Mock()
    fake_attachments.Item.side_effect = RuntimeError("attachment index out of range")

    fake_mail_item = Mock()
    fake_mail_item.Attachments = fake_attachments

    namespace = Mock()
    namespace.GetItemFromID.return_value = fake_mail_item

    session_factory = Mock()
    session_factory.session.return_value.__enter__ = Mock(return_value=namespace)
    session_factory.session.return_value.__exit__ = Mock(return_value=False)

    message = MailMessage(
        entry_id="entry-001",
        store_id="store-001",
        subject="每日庫存",
        sender_name="Warehouse",
        sender_email="warehouse@example.com",
        received_time=datetime(2026, 9, 24, 8, 0),
        unread=False,
        attachments=(
            AttachmentInfo(
                index=1,
                filename="每日庫存.xlsx",
                extension=".xlsx",
            ),
        ),
    )

    client = OutlookClient(session_factory=session_factory)

    with pytest.raises(MailAccessError, match="index 1"):
        client.download_attachments(
            message,
            output_dir=tmp_path,
            extensions=(".xlsx",),
        )

    fake_attachments.Item.assert_called_once_with(1)


def test_download_latest_raises_when_no_message_matches(
    tmp_path: Path,
) -> None:
    client = OutlookClient()

    client.find_latest = Mock(return_value=None)  # type: ignore[method-assign]

    with pytest.raises(MailNotFoundError):
        client.download_latest(
            MailQuery(subject_contains="每日庫存"),
            output_dir=tmp_path,
        )


def test_download_latest_downloads_matching_message(
    tmp_path: Path,
) -> None:

    message = MailMessage(
        entry_id="entry-001",
        store_id="store-001",
        subject="每日庫存",
        sender_name="Warehouse",
        sender_email="warehouse@example.com",
        received_time=datetime(2026, 9, 24, 8, 0),
        unread=False,
        attachments=(
            AttachmentInfo(
                index=1,
                filename="每日庫存.xlsx",
                extension=".xlsx",
            ),
        ),
    )

    query = MailQuery(
        subject_contains="每日庫存",
        attachment_name_contains="庫存",
        attachment_extensions=(".xlsx",),
    )

    client = OutlookClient()

    client.find_latest = Mock(return_value=message)  # type: ignore[method-assign]
    client.download_attachments = Mock(return_value=[tmp_path / "每日庫存.xlsx"])  # type: ignore[method-assign]

    result = client.download_latest(
        query,
        output_dir=tmp_path,
        conflict="rename",
    )

    assert result == [tmp_path / "每日庫存.xlsx"]

    client.find_latest.assert_called_once_with(query)

    client.download_attachments.assert_called_once_with(
        message,
        tmp_path,
        extensions=query.attachment_extensions,
        filename=query.attachment_name,
        filename_contains=query.attachment_name_contains,
        conflict="rename",
    )


def test_search_recursive_merges_and_sorts_across_subfolders(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class FakeMailItem:
        Class = 43

    root_items = FakeItems([FakeMailItem()])
    child_items = FakeItems([FakeMailItem()])

    child_folder = FakeFolder(child_items)
    root_folder = FakeFolder(root_items, folders=[child_folder])

    namespace = object()

    older_message = make_message("entry-old", received_time=datetime(2026, 9, 20, 8, 0))
    newer_message = make_message("entry-new", received_time=datetime(2026, 9, 23, 8, 0))
    messages = iter([older_message, newer_message])

    monkeypatch.setattr(client_module, "resolve_folder", lambda namespace, path: root_folder)
    monkeypatch.setattr(client_module, "build_restrict_filter", lambda query: None)
    monkeypatch.setattr(
        client_module, "mail_item_to_message", lambda item, store_id: next(messages)
    )
    monkeypatch.setattr(client_module, "message_matches", lambda message, query: True)

    client = OutlookClient(session_factory=FakeSessionFactory(namespace))

    result = client.search(
        MailQuery(recursive=True),
    )

    assert [message.entry_id for message in result] == ["entry-new", "entry-old"]


def test_search_recursive_applies_limit_after_global_sort(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class FakeMailItem:
        Class = 43

    root_items = FakeItems([FakeMailItem()])
    child_items = FakeItems([FakeMailItem()])

    child_folder = FakeFolder(child_items)
    root_folder = FakeFolder(root_items, folders=[child_folder])

    namespace = object()

    older_message = make_message("entry-old", received_time=datetime(2026, 9, 20, 8, 0))
    newer_message = make_message("entry-new", received_time=datetime(2026, 9, 23, 8, 0))
    messages = iter([older_message, newer_message])

    monkeypatch.setattr(client_module, "resolve_folder", lambda namespace, path: root_folder)
    monkeypatch.setattr(client_module, "build_restrict_filter", lambda query: None)
    monkeypatch.setattr(
        client_module, "mail_item_to_message", lambda item, store_id: next(messages)
    )
    monkeypatch.setattr(client_module, "message_matches", lambda message, query: True)

    client = OutlookClient(session_factory=FakeSessionFactory(namespace))

    result = client.search(MailQuery(recursive=True), limit=1)

    assert [message.entry_id for message in result] == ["entry-new"]


def test_search_recursive_skips_subfolder_that_fails_to_read(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class FakeMailItem:
        Class = 43

    class BrokenItems:
        def Sort(self, property_name: str, descending: bool) -> None:
            pass

        def __iter__(self) -> Iterator[Any]:
            raise RuntimeError("access denied")

    class BrokenFolder:
        def __init__(self) -> None:
            self.Items = BrokenItems()
            self.Folders: list[Any] = []
            self.StoreID = "store-broken"

    root_items = FakeItems([FakeMailItem()])
    root_folder = FakeFolder(root_items, folders=[BrokenFolder()])

    namespace = object()
    message = make_message("entry-1")

    monkeypatch.setattr(client_module, "resolve_folder", lambda namespace, path: root_folder)
    monkeypatch.setattr(client_module, "build_restrict_filter", lambda query: None)
    monkeypatch.setattr(client_module, "mail_item_to_message", lambda item, store_id: message)
    monkeypatch.setattr(client_module, "message_matches", lambda message, query: True)

    client = OutlookClient(session_factory=FakeSessionFactory(namespace))

    result = client.search(MailQuery(recursive=True))

    assert result == [message]
