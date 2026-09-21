class OutlookError(Exception):
    """Base exception for this package."""


class OutlookUnavailableError(OutlookError):
    """Outlook Desktop or pywin32 is unavailable."""


class OutlookConnectionError(OutlookError):
    """A COM session could not be opened."""


class FolderNotFoundError(OutlookError):
    """The requested Outlook folder path does not exist."""


class MailNotFoundError(OutlookError):
    """No message matched a high-level download request."""


class MailAccessError(OutlookError):
    """A previously found message can no longer be opened."""


class AttachmentNotFoundError(OutlookError):
    """No attachment matched the requested download filters."""


class AttachmentConflictError(OutlookError):
    """A destination file exists and conflict policy is error."""


class AttachmentSaveError(OutlookError):
    """Outlook failed to save an attachment."""