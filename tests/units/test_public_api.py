from email_downloader import (
    AttachmentInfo,
    AttachmentNotFoundError,
    FolderNotFoundError,
    MailMessage,
    MailNotFoundError,
    MailQuery,
    OutlookClient,
    OutlookError,
)


def test_public_api_is_importable() -> None:
    assert OutlookClient is not None
    assert MailQuery is not None
    assert MailMessage is not None
    assert AttachmentInfo is not None
    assert OutlookError is not None
    assert FolderNotFoundError is not None
    assert MailNotFoundError is not None
    assert AttachmentNotFoundError is not None
