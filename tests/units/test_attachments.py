from pathlib import Path

import pytest

from email_downloader.attachments import (
    resolve_destination,
    safe_attachment_name,
    select_attachments,
)
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


@pytest.mark.parametrize(
    ("filename", "expected"),
    [
        ("庫存表.xlsx", "庫存表.xlsx"),
        (r"..\庫存表.xlsx", "庫存表.xlsx"),
        ("../庫存表.xlsx", "庫存表.xlsx"),
        (r"C:\Temp\庫存表.xlsx", "庫存表.xlsx"),
    ],
)

def test_safe_attachment_name_returns_basename(
    filename: str,
    expected: str,
) -> None:
    assert safe_attachment_name(filename) == expected


@pytest.mark.parametrize(
    "filename",
    [
        "",
        ".",
        "..",
    ],
)

def test_safe_attachment_name_rejects_invalid_names(filename: str) -> None:
    with pytest.raises(ValueError):
        safe_attachment_name(filename)


def test_safe_attachment_destination_stays_inside_output_dir(tmp_path: Path) -> None:
    filename = r"..\..\outside.xlsx"

    safe_name = safe_attachment_name(filename)
    destination = tmp_path / safe_name

    assert destination.resolve().parent == tmp_path.resolve()


def test_resolve_destination_rename_preserves_multiple_suffixes(tmp_path: Path) -> None:
    destination = tmp_path / "inventory.tar.gz"
    destination.touch()

    result = resolve_destination(destination, conflict="rename")

    assert result == tmp_path / "inventory_1.tar.gz"

@pytest.mark.parametrize(
    ("filename", "expected"),
    [
        ("庫存表.xlsx", "庫存表.xlsx"),
        (r"..\庫存表.xlsx", "庫存表.xlsx"),
        ("../庫存表.xlsx", "庫存表.xlsx"),
        (r"C:\Temp\庫存表.xlsx", "庫存表.xlsx"),
    ],
)

def test_safe_attachment_name_returns_basename(
    filename: str,
    expected: str,
) -> None:
    assert safe_attachment_name(filename) == expected


@pytest.mark.parametrize(
    "filename",
    [
        "",
        ".",
        "..",
    ],
)

def test_safe_attachment_name_rejects_invalid_names(filename: str) -> None:
    with pytest.raises(ValueError):
        safe_attachment_name(filename)


def test_safe_attachment_destination_stays_inside_output_dir(tmp_path: Path) -> None:
    filename = r"..\..\outside.xlsx"

    safe_name = safe_attachment_name(filename)
    destination = tmp_path / safe_name

    assert destination.resolve().parent == tmp_path.resolve()


def test_resolve_destination_rename_preserves_multiple_suffixes(tmp_path: Path) -> None:
    destination = tmp_path / "inventory.tar.gz"
    destination.touch()

    result = resolve_destination(destination, conflict="rename")

    assert result == tmp_path / "inventory_1.tar.gz"