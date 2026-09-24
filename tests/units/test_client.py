from contextlib import contextmanager
from typing import Any

import pytest

from datetime import datetime

from email_downloader.models import MailQuery
import email_downloader.client as client_module

from email_downloader.client import OutlookClient
from email_downloader.models import MailMessage


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

    def __iter__(self):
        return iter(self._items)


class FakeFolder:
    def __init__(self, items: FakeItems) -> None:
        self.Items = items
        self.StoreID = "store-123"

class FakeSessionFactory:
    def __init__(self, namespace: object) -> None:
        self.namespace = namespace

    @contextmanager
    def session(self):
        yield self.namespace


def make_message(
    entry_id: str,
    *,
    subject: str = "Daily Inventory",
) -> MailMessage:
    return MailMessage(
        entry_id=entry_id,
        store_id="store-123",
        subject=subject,
        sender_name="Warehouse",
        sender_email="warehouse@example.com",
        received_time=datetime(2026, 9, 23, 8, 0),
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

    client = OutlookClient(
        session_factory=FakeSessionFactory(namespace)
    )

    result = client.search(
        MailQuery(has_attachment=True)
    )

    assert result == [message]

    assert items.restrict_calls == [
        "[HasAttachment] = True"
    ]

    assert items.sort_calls == [
        ("[ReceivedTime]", True)
    ]

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

    client = OutlookClient(
        session_factory=FakeSessionFactory(namespace)
    )

    result = client.search(
        MailQuery()
    )

    assert result == []
    assert items.restrict_calls == []

    assert items.sort_calls == [
        ("[ReceivedTime]", True)
    ]

@pytest.mark.parametrize("limit", [0, -1])
def test_search_rejects_invalid_limit(limit: int) -> None:
    client = OutlookClient(
        session_factory=FakeSessionFactory(object())
    )

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

    client = OutlookClient(
        session_factory=FakeSessionFactory(namespace)
    )

    result = client.search(
        MailQuery()
    )

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

    client = OutlookClient(
        session_factory=FakeSessionFactory(namespace)
    )

    result = client.search(
        MailQuery()
    )

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

    client = OutlookClient(
        session_factory=FakeSessionFactory(namespace)
    )

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

    client = OutlookClient(
        session_factory=FakeSessionFactory(namespace)
    )

    result = client.search(
        MailQuery()
    )

    assert result == [message]


def test_find_latest_returns_first_search_result(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    client = OutlookClient(
        session_factory=FakeSessionFactory(object())
    )

    message = make_message("entry-1")

    monkeypatch.setattr(
        client,
        "search",
        lambda query, *, limit=None: [message],
    )

    result = client.find_latest(
        MailQuery()
    )

    assert result == message


def test_find_latest_returns_none_when_no_message_matches(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    client = OutlookClient(
        session_factory=FakeSessionFactory(object())
    )

    monkeypatch.setattr(
        client,
        "search",
        lambda query, *, limit=None: [],
    )

    result = client.find_latest(
        MailQuery()
    )

    assert result is None