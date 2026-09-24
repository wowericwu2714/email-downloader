from email_downloader.com_backend import (
    Pywin32SessionFactory,
    mail_item_to_message,
)
from email_downloader.filters import (
    build_restrict_filter,
    message_matches,
)
from email_downloader.folders import resolve_folder
from email_downloader.models import MailMessage, MailQuery
from email_downloader.protocols import ComSessionFactory


class OutlookClient:
    """High-level Outlook mail search client."""

    def __init__(
        self, 
        session_factory: ComSessionFactory | None = None,
    ) -> None:
        self._session_factory = session_factory or Pywin32SessionFactory()

    def search(
        self, 
        query: MailQuery,
        *, 
        limit: int | None = None,
    ) -> list[MailMessage]:
        """Search Outlook messages matching the query."""
        if limit is not None and limit < 1:
            raise ValueError("limit must be at least 1")

        results: list[MailMessage] = []

        with self._session_factory.session() as namespace:
            folder = resolve_folder(namespace, query.folder)
            items = folder.Items

            restrict_filter = build_restrict_filter(query)

            if restrict_filter is not None:
                items = items.Restrict(restrict_filter)

            items.Sort("[ReceivedTime]", True)

            for item in items:
                if getattr(item, "Class", None) != 43:  # 43 corresponds to MailItem
                    continue

                try:
                    message = mail_item_to_message(
                        item,
                        folder.StoreID,
                    )
                except Exception:  # noqa: BLE001, S112 -- skip unreadable item; logging deferred to a later task.
                    continue

                if not message_matches(message, query):
                    continue

                results.append(message)

                if limit is not None and len(results) >= limit:
                    break

        return results

    def find_latest(
        self,
        query: MailQuery,
    ) -> MailMessage | None:
        """Return the newest message matching the query."""
        messages = self.search(query, limit=1)
        return messages[0] if messages else None