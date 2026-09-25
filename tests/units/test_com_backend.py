import sys
from datetime import datetime
from types import ModuleType

import pytest

from email_downloader.com_backend import (
    Pywin32SessionFactory,
    get_sender_email,
    mail_item_to_message,
)
from email_downloader.exceptions import OutlookConnectionError


def install_fake_pywin32(
    monkeypatch: pytest.MonkeyPatch,
    events: list[str],
    *,
    dispatch_fails: bool = False,
) -> object:
    pythoncom = ModuleType("pythoncom")
    win32com = ModuleType("win32com")
    win32com_client = ModuleType("win32com.client")

    namespace = object()

    def co_initialize() -> None:
        events.append("CoInitialize")

    def co_uninitialize() -> None:
        events.append("CoUninitialize")

    class FakeApplication:
        def GetNamespace(self, name: str) -> object:
            events.append(f"GetNamespace({name})")
            return namespace

    def dispatch(name: str) -> FakeApplication:
        events.append(f"Dispatch({name})")

        if dispatch_fails:
            raise RuntimeError("Outlook unavailable")

        return FakeApplication()

    pythoncom.CoInitialize = co_initialize      # type: ignore[attr-defined]
    pythoncom.CoUninitialize = co_uninitialize  # type: ignore[attr-defined]
    win32com_client.Dispatch = dispatch        # type: ignore[attr-defined]
    win32com.client = win32com_client          # type: ignore[attr-defined]

    monkeypatch.setitem(sys.modules, "pythoncom", pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", win32com_client)

    return namespace


def test_session_initializes_and_uninitializes_com(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    events: list[str] = []
    expected_namespace = install_fake_pywin32(monkeypatch, events)

    factory = Pywin32SessionFactory()

    with factory.session() as namespace:
        assert namespace is expected_namespace
        assert events == [
            "CoInitialize",
            "Dispatch(Outlook.Application)",
            "GetNamespace(MAPI)",
        ]

    assert events == [
        "CoInitialize",
        "Dispatch(Outlook.Application)",
        "GetNamespace(MAPI)",
        "CoUninitialize",
    ]


def test_session_uninitializes_com_when_dispatch_fails(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    events: list[str] = []
    install_fake_pywin32(monkeypatch, events, dispatch_fails=True)

    factory = Pywin32SessionFactory()

    with pytest.raises(
        OutlookConnectionError,
        match="Could not open Outlook MAPI session",
    ), factory.session():
        pass

    assert events == [
        "CoInitialize",
        "Dispatch(Outlook.Application)",
        "CoUninitialize",
    ]


def test_get_sender_email_returns_smtp_address() -> None:
    class FakeItem:
        SenderEmailType = "SMTP"
        SenderEmailAddress = "warehouse@example.com"
    
    assert get_sender_email(FakeItem()) == "warehouse@example.com"


def test_get_sender_email_resolves_exchange_primary_smtp() -> None:
    class FakeExchangeUser:
        PrimarySmtpAddress = "warehouse@example.com"

    class FakeSender:
        def GetExchangeUser(self) -> FakeExchangeUser:
            return FakeExchangeUser()

    class FakeItem:
        SenderEmailType = "EX"
        SenderEmailAddress = "/O=COMPANY/OU=EXCHANGE/CN=RECIPIENTS/CN=WAREHOUSE"
        Sender = FakeSender()
       
    assert get_sender_email(FakeItem()) == "warehouse@example.com"


def test_get_sender_email_falls_back_when_exchange_resolution_fails() -> None:
    original_address = "/O=COMPANY/OU=EXCHANGE/CN=RECIPIENTS/CN=WAREHOUSE"

    class FakeSender:
        def GetExchangeUser(self) -> object:
            raise RuntimeError("Exchange lookup failed")

    class FakeItem:
        SenderEmailType = "EX"
        SenderEmailAddress = original_address
        Sender = FakeSender()

    assert get_sender_email(FakeItem()) == original_address


class FakeAttachment:
    def __init__(self, filename: str, size: int | None = None) -> None:
        self.Filename = filename
        if size is not None:
            self.Size = size
    

class FakeAttachments:
    def __init__(self, *attachments: FakeAttachment) -> None:
        self._attachments = attachments
        self.Count = len(attachments)

    def Item(self, index: int) -> FakeAttachment:
        return self._attachments[index - 1]

def test_mail_item_to_message_maps_fields_and_attachments() -> None:
    received_time = datetime(2026, 9, 22, 10, 30)

    class FakeItem:
        Class = 43
        EntryID = "entry-123"
        Subject = "Daily Inventory"
        SenderName = "Warehouse"
        SenderEmailType = "SMTP"
        SenderEmailAddress = "warehouse@example.com"
        ReceivedTime = received_time
        UnRead = True
        Attachments = FakeAttachments(
            FakeAttachment("Daily Inventory.XLSX", 12345),
            FakeAttachment("note.pdf"),
        )

    message = mail_item_to_message(FakeItem(), "store-456")

    assert message.entry_id == "entry-123"
    assert message.store_id == "store-456"
    assert message.subject == "Daily Inventory"
    assert message.sender_name == "Warehouse"
    assert message.sender_email == "warehouse@example.com"
    assert message.received_time == received_time
    assert message.unread is True

    assert len(message.attachments) == 2

    assert message.attachments[0].index == 1
    assert message.attachments[0].filename == "Daily Inventory.XLSX"
    assert message.attachments[0].extension == ".xlsx"
    assert message.attachments[0].size == 12345

    assert message.attachments[1].index == 2
    assert message.attachments[1].filename == "note.pdf"
    assert message.attachments[1].extension == ".pdf"
    assert message.attachments[1].size is None


def test_mail_item_to_message_rejects_non_mail_item() -> None:
    class FakeItem:
        Class = 99

    with pytest.raises(ValueError):
        mail_item_to_message(FakeItem(), "store-456") 


def test_mail_item_to_message_requires_entry_id() -> None:
    class FakeItem:
        Class = 43
        Subject = "Daily Inventory"
        ReceivedTime = datetime(2026, 9, 22, 10, 30)
        SenderName = "Warehouse"
        SenderEmailType = "SMTP"
        SenderEmailAddress = "warehouse@example.com"
        UnRead = False
        Attachments = FakeAttachments()

    with pytest.raises(AttributeError):
        mail_item_to_message(FakeItem(), "store-456")


def test_mail_item_to_message_handles_missing_optional_fields() -> None:
    class FakeItem:
        Class = 43
        EntryID = "entry-123"
        Subject = "Daily Inventory"
        ReceivedTime = datetime(2026, 9, 22, 10, 30)
        Attachments = FakeAttachments()

    message = mail_item_to_message(FakeItem(), "store-456")

    assert message.sender_name == ""
    assert message.sender_email is None
    assert message.unread is False
    assert message.attachments == ()