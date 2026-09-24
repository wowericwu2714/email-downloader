from pathlib import Path

import pytest


from email_downloader.attachments import resolve_destination, select_attachments
from email_downloader.exceptions import AttachmentConflictError
from email_downloader.models import AttachmentInfo


def test_selects_attachment_with_combined_filters() -> None:
    attachments = (
        AttachmentInfo(1, "每日庫存.pdf", ".pdf"),
        AttachmentInfo(2, "每日庫存.XLSX", ".xlsx"),
    )

    result = select_attachments(
        attachments,
        filename_contains="庫存",
        extensions=("xlsx",),
    )

    assert [item.index for item in result] == [2]


def test_resolve_destination_overwrite_returns_original_path(tmp_path: Path) -> None:
    destination = tmp_path / "庫存表.xlsx"
    destination.touch()

    result = resolve_destination(destination, conflict="overwrite")

    assert result == destination

def test_resolve_destination_skip_returns_none(tmp_path: Path) -> None:
    destination = tmp_path / "庫存表.xlsx"
    destination.touch()

    result = resolve_destination(destination, conflict="skip")

    assert result is None

def test_resolve_destination_error_raises(tmp_path: Path) -> None:
    destination = tmp_path / "庫存表.xlsx"
    destination.touch()

    with pytest.raises(AttachmentConflictError):
        resolve_destination(destination, conflict="error")

def test_resolve_destination_rename_uses_next_available_name(tmp_path: Path) -> None:
    destination = tmp_path / "庫存表.xlsx"
    destination.touch()
    (tmp_path / "庫存表_1.xlsx").touch()

    result = resolve_destination(destination, conflict="rename")

    assert result == tmp_path / "庫存表_2.xlsx"
