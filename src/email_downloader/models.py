from dataclasses import dataclass
from datetime import datetime
from typing import Literal, TypeAlias

ConflictPolicy: TypeAlias = Literal["overwrite", "skip", "rename", "error"]

@dataclass(frozen=True, slots=True)
class AttachmentInfo:
    index: int
    filename: str
    extension: str
    size: int | None = None

@dataclass(frozen=True, slots=True)
class MailMessage:
    entry_id: str
    store_id: str
    subject: str
    sender_name: str
    sender_email: str | None
    received_time: datetime
    unread: bool
    attachments: tuple[AttachmentInfo, ...] # ...表示長度不限

@dataclass(frozen=True, slots=True)
class MailQuery:
    folder: str = "收件匣"
    sender: str | None = None
    sender_contains: str | None = None
    subject: str | None = None
    subject_contains: str | None = None
    received_after: datetime | None = None
    received_before: datetime | None = None
    has_attachment: bool | None = None
    attachment_name: str | None = None
    attachment_name_contains: str | None = None
    attachment_extensions: tuple[str, ...] = ()
    unread_only: bool = False
    recursive: bool = False

    def __post_init__(self) -> None:
        # Folder 不可為空
        if not self.folder.strip():
            raise ValueError("folder must not be empty")

        # 時間必須是不帶時區的本機 datetime
        for field_name, value in (
            ("received_after", self.received_after),
            ("received_before", self.received_before),
        ):
            if value is not None and value.utcoffset() is not None:
                raise ValueError(f"{field_name} must be a naive local datetime")

        # 開始時間不可晚於結束時間
        if (
            self.received_after is not None
            and self.received_before is not None
            and self.received_after > self.received_before
        ):
            raise ValueError("received_after must be earlier than received_before")

        # exact 與 contains 不可同時使用
        exclusive_fields = (
            ("sender", "sender_contains"),
            ("subject", "subject_contains"),
            ("attachment_name", "attachment_name_contains"),
        )

        for exact_field, contains_field in exclusive_fields:
            if (
                getattr(self, exact_field) is not None
                and getattr(self, contains_field) is not None
            ):
                raise ValueError(
                    f"{exact_field} and {contains_field} cannot be used simultaneously"
                )

        object.__setattr__(self, "attachment_extensions", normalize_extensions(self.attachment_extensions))


def normalize_extensions(extensions: tuple[str, ...]) -> tuple[str, ...]:
    """Normalize attachment extensions to lowercase, leading-dot form, deduped."""
    return tuple(
        dict.fromkeys(
            f".{ext.strip().lstrip('.').casefold()}"
            for ext in extensions
        )
    )