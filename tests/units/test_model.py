from datetime import datetime

import pytest

from email_downloader.models import MailQuery


def test_query_normalizes_extensions() -> None:
    query = MailQuery(attachment_extensions=("XLSX", ".xls", ".XLSX"))
    assert query.attachment_extensions == (".xlsx", ".xls")

def test_query_rejects_reversed_time_range() -> None:
    with pytest.raises(ValueError, match="received_after"):
        MailQuery(
            received_after=datetime(2026, 9, 22), 
            received_before=datetime(2026, 9, 21)
        )


def test_query_rejects_timezone_aware_datetime() -> None:
    aware = datetime.fromisoformat("2026-09-21T08:00:00+08:00")
    with pytest.raises(ValueError, match="naive local datetime"):
        MailQuery(
            received_after=aware
        )

def test_query_rejects_exact_and_contains_for_same_field() -> None:
    with pytest.raises(ValueError, match="sender"):
        MailQuery(
            sender="a@example.com",
            sender_contains="example"
        )