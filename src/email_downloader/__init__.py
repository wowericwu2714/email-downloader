from .client import OutlookClient
from .exceptions import (
    AttachmentConflictError,
    AttachmentNotFoundError,
    AttachmentSaveError,
    FolderNotFoundError,
    MailAccessError,
    MailNotFoundError,
    OutlookConnectionError,
    OutlookError,
    OutlookUnavailableError,
)
from .models import AttachmentInfo, MailMessage, MailQuery

__version__ = "0.1.0"

__all__ = [
    "AttachmentConflictError",
    "AttachmentInfo",
    "AttachmentNotFoundError",
    "AttachmentSaveError",
    "FolderNotFoundError",
    "MailAccessError",
    "MailMessage",
    "MailNotFoundError",
    "MailQuery",
    "OutlookClient",
    "OutlookConnectionError",
    "OutlookError",
    "OutlookUnavailableError",
]
