import os
from datetime import datetime, timedelta
from pathlib import Path

import pytest

from email_downloader import MailQuery, OutlookClient

pytestmark = pytest.mark.outlook_integration

@pytest.mark.skipif(
    os.getenv("RUN_OUTLOOK_INTEGRATION") != "1",
    reason="set RUN_OUTLOOK_INTEGRATION=1 on a Windows Outlook workstation",
)  
def test_can_search_recent_inbox_mail() -> None:
    client = OutlookClient()
    messages = client.search(
        MailQuery(received_after=datetime.now() - timedelta(days=1)), 
        limit=1,
    )
    assert len(messages) <= 1


@pytest.mark.skipif(
    os.getenv("RUN_OUTLOOK_INTEGRATION") != "1",
    reason="set RUN_OUTLOOK_INTEGRATION=1 on a Windows Outlook workstation",
)
def test_can_find_and_optionally_download_test_message(tmp_path: Path) -> None:
    """Find a dedicated smoke-test message and optionally download its attachments.

    Point OUTLOOK_TEST_FOLDER / OUTLOOK_TEST_SUBJECT_CONTAINS at a
    disposable Outlook folder/message set up specifically for this
    test. This test never moves, deletes, or marks the message read.
    """
    query = MailQuery(
        folder=os.getenv("OUTLOOK_TEST_FOLDER", "收件匣"),
        subject_contains=os.getenv("OUTLOOK_TEST_SUBJECT_CONTAINS"),
    )

    client = OutlookClient()
    message = client.find_latest(query)

    if message is None:
        pytest.skip("No message matched OUTLOOK_TEST_FOLDER / OUTLOOK_TEST_SUBJECT_CONTAINS")

    print(f"\nsubject={message.subject!r}")
    print(f"sender={message.sender_email!r}")
    print(f"attachments={[a.filename for a in message.attachments]!r}")

    if os.getenv("RUN_OUTLOOK_DOWNLOAD_INTEGRATION") != "1":
        pytest.skip("set RUN_OUTLOOK_DOWNLOAD_INTEGRATION=1 to also test download_attachments()")

    paths = client.download_attachments(message, output_dir=tmp_path, conflict="skip")
    assert all(path.exists() for path in paths)