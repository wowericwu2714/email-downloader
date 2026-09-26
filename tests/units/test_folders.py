from __future__ import annotations

from collections.abc import Iterator

import pytest

from email_downloader.exceptions import FolderNotFoundError
from email_downloader.folders import iter_folder_tree, resolve_folder, split_folder_path


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        ("收件匣/外倉/MSS", ("外倉", "MSS")),
        (r"收件匣\外倉\MSS", ("外倉", "MSS")),
        ("Inbox/Shipping Schedule", ("Shipping Schedule",)),
        ("外倉/MSS", ("外倉", "MSS")),
    ]
)

def test_split_folder_path(
    value: str, 
    expected: tuple[str, ...],
) -> None:
    assert split_folder_path(value) == expected

def test_split_folder_path_rejects_empty_segment() -> None:
    with pytest.raises(
        ValueError,
        match="empty segment"
    ):
        split_folder_path("收件匣//MSS")

class FakeFolders:
    def __init__(self, folders: dict[str, FakeFolder]) -> None:
        self._folders = folders

    def Item(self, name: str) -> FakeFolder:
        return self._folders[name]

    def __iter__(self) -> Iterator[FakeFolder]:
        return iter(self._folders.values())

class FakeFolder:
    def __init__(self, name: str, children: dict[str, FakeFolder] | None = None, ) -> None:
        self.name = name
        self.Folders = FakeFolders(children or {})


class FakeNamespace:
    def __init__(self, inbox: FakeFolder) -> None:
        self._inbox = inbox

    def GetDefaultFolder(self, folder_type: int) -> FakeFolder:
        if folder_type != 6:
            raise ValueError(f"Unsupported folder type: {folder_type}")

        return self._inbox

def test_resolve_folder_returns_nested_folder() -> None:
    mss = FakeFolder("MSS")
    warehouse = FakeFolder("外倉", children={"MSS": mss})
    inbox = FakeFolder("收件匣", children={"外倉": warehouse})
    namespace = FakeNamespace(inbox)

    resolved = resolve_folder(namespace, "收件匣/外倉/MSS")
    assert resolved is mss

def test_resolve_folder_raises_package_error_for_missing_folder() -> None:
    inbox = FakeFolder("收件匣")
    namespace = FakeNamespace(inbox)
    path = "收件匣/不存在"

    with pytest.raises(FolderNotFoundError) as exc_info:
        resolve_folder(namespace, path)

    assert path in str(exc_info.value)

def test_iter_folder_tree_yields_folder_and_all_descendants() -> None:
    grandchild = FakeFolder("孫")
    child_a = FakeFolder("A", children={"孫": grandchild})
    child_b = FakeFolder("B")
    root = FakeFolder("根", children={"A": child_a, "B": child_b})

    result = list(iter_folder_tree(root))

    assert result == [root, child_a, grandchild, child_b]