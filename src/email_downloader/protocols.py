from contextlib import AbstractContextManager
from typing import Any, Protocol


class ComSessionFactory(Protocol):
    """Factory for Outlook COM sessions."""

    def session(self) -> AbstractContextManager[Any]:
        """Return a context manager yielding an Outlook MAPI namespace."""
        ...
