# Email Downloader (Outlook Client) Package Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use `superpowers:subagent-driven-development` (recommended) or `superpowers:executing-plans` to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 建立一個可被 ICS、外倉庫存、Shipping Schedule 等專案直接引入的 Windows Outlook Desktop 搜尋與附件下載 Python package。

**Architecture:** 對外只暴露純 Python 的 `OutlookClient`、`MailQuery`、`MailMessage` 與例外類別；Outlook COM、資料夾定位、查詢縮小、信件轉換及附件落地都封裝在 package 內。搜尋結果只保留 `EntryID`／`StoreID` 等定位資訊，不把 COM object 傳給呼叫端；下載時再於同一個 COM session 內重新取得信件。

**Tech Stack:** Python 3.11+、uv、pywin32、pytest、pytest-cov、ruff、mypy、Windows 11、Classic Outlook Desktop COM。

**Spec:** 本文件的「設計基準」章節；內容依 2026-09-21 已確認的需求與架構建議整理。

## Global Constraints

- 僅支援 Windows 11 與 Classic Outlook Desktop；New Outlook、macOS Outlook、Microsoft Graph、IMAP 不在 v1 範圍。
- package 名稱為 `outlook-client`，Python import 名稱為 `outlook_client`。
- package 只處理「如何找信、如何下載」；MSS、裕國、Shipping Schedule 等業務條件留在呼叫端。
- 對外不得回傳 `win32com` COM object。
- `win32com` 與 `pythoncom` 必須 lazy import，使非 Windows 環境仍可 import models 並執行純函式測試。
- 每次 COM 操作都必須在目前執行緒執行 `CoInitialize()`／`CoUninitialize()`，以支援 Prefect、APScheduler 或其他 worker thread。
- package logger 只使用 `logging.getLogger(__name__)`，不得呼叫 `logging.basicConfig()`。
- 所有路徑 API 接受 `str | os.PathLike[str]`，內部統一轉為 `pathlib.Path`。
- 所有副檔名比較不分大小寫，且統一為含前導點的形式，例如 `.xlsx`。
- 所有 exact／contains 字串比對預設不分大小寫；前後空白保留，除非欄位驗證明確正規化。
- v1 的 `received_after`／`received_before` 接受 Outlook 本機時間的 naive `datetime`；傳入 timezone-aware `datetime` 時丟出 `ValueError`，避免無聲時區偏移。
- v1 原本不支援遞迴搜尋所有子資料夾；此限制已由 Task 12（`MailQuery.recursive`）解除，詳見該 Task。v1 仍不支援寫信、回信、搬移、刪除或改成已讀。
- 所有公開 API 都要有 type hints 與 docstring。
- 採 TDD：每個功能先有失敗測試，再寫最小實作，再跑完整測試。

## Review Focus

以下五項必須各有測試覆蓋：

1. 排程器 worker thread 中呼叫 COM：必須初始化並在結束時釋放該執行緒的 COM apartment。
2. Exchange 內部寄件者回傳 `EX` address：應優先解析 `PrimarySmtpAddress`，失敗時安全退回原始地址。
3. Outlook 資料夾含中文、反斜線或斜線：`收件匣/外倉/MSS` 與 `收件匣\外倉\MSS` 應解析為相同路徑。
4. 同名附件競爭：`overwrite`、`skip`、`rename`、`error` 四種策略都必須有確定且可測試的結果。
5. 搜尋結果取得後信件被移動或刪除：下載時應轉成 package 自訂的 `MailAccessError`，不得洩漏 pywin32 例外。

---

## 設計基準

### 1. Public API

其他專案最終只需要這樣使用：

```python
from datetime import datetime, timedelta
from outlook_client import MailQuery, OutlookClient

client = OutlookClient()

files = client.download_latest(
    query=MailQuery(
        folder="收件匣/外倉/MSS",
        sender_contains="MSS",
        subject_contains="每日庫存",
        received_after=datetime.now() - timedelta(days=2),
        attachment_name_contains="庫存",
        attachment_extensions=(".xlsx", ".xls"),
    ),
    output_dir=r"D:\data\warehouse",
    conflict="rename",
)
```

低階 API 保留給需要先檢查信件再決定是否下載的專案：

```python
messages = client.search(query, limit=20)
latest = client.find_latest(query)

if latest is not None:
    files = client.download_attachments(
        latest,
        output_dir=r"D:\data\warehouse",
        extensions=(".xlsx",),
        filename_contains="庫存",
        conflict="skip",
    )
```

### 2. 搜尋資料流

```text
MailQuery
  -> 解析指定 Outlook folder
  -> Outlook Items.Restrict 縮小日期／已讀／附件候選集
  -> ReceivedTime 由新到舊排序
  -> Python 比對 sender／subject／附件檔名／副檔名
  -> 轉成純 Python MailMessage
```

`Items.Restrict` 用於大量信件中的候選縮小；Microsoft 文件說明它會回傳符合 filter 的新 collection，且在大型 collection、少量命中時通常較快。contains 條件與附件名稱則在 Python 完成，避免把複雜 DASL 語法暴露給使用者。

### 3. 信件與附件定位

- `MailMessage.entry_id` 與 `MailMessage.store_id` 是重新取得 Outlook item 的定位資訊。
- `MailMessage` 仍是純 dataclass，不保留 COM reference。
- `download_attachments()` 開啟新的 COM session，使用 `NameSpace.GetItemFromID(entry_id, store_id)` 重新取得信件。
- COM session 結束後不可再存取任何 COM object。

### 4. 找不到信件的行為

- `search()`：回傳空 list。
- `find_latest()`：回傳 `None`。
- `download_latest()`：丟出 `MailNotFoundError`。
- 符合信件但沒有符合下載條件的附件：丟出 `AttachmentNotFoundError`。

這讓呼叫端能明確處理「今日沒收到外倉庫存表，沿用最近結存」等業務流程。

### 5. v1 不納入的項目

- YAML／JSON 設定驅動下載。
- CLI 指令。
- Microsoft Graph、IMAP 或 macOS backend。
- 跨 mailbox／shared mailbox 的頂層 store 選擇。
- 本文搜尋、收件者、CC、分類、重要性。
- 下載紀錄資料庫、hash 去重與排程功能。
- Outlook 信件狀態修改。

---

## 目標檔案結構

```text
outlook-client/
├─ pyproject.toml
├─ README.md
├─ LICENSE
├─ src/
│  └─ outlook_client/
│     ├─ __init__.py          # 穩定 public exports
│     ├─ client.py            # OutlookClient facade 與 use cases
│     ├─ models.py            # MailQuery、MailMessage、AttachmentInfo
│     ├─ filters.py           # Restrict 字串與純 Python matcher
│     ├─ folders.py           # Inbox-based folder path resolver
│     ├─ attachments.py       # 檔名篩選與 collision policy
│     ├─ com_backend.py       # lazy pywin32、COM session、COM item mapping
│     ├─ exceptions.py        # package 自訂例外
│     ├─ protocols.py         # 測試注入所需 Protocol
│     └─ py.typed             # PEP 561 marker
└─ tests/
   ├─ unit/
   │  ├─ test_models.py
   │  ├─ test_filters.py
   │  ├─ test_folders.py
   │  ├─ test_attachments.py
   │  ├─ test_com_backend.py
   │  └─ test_client.py
   └─ integration/
      └─ test_outlook_smoke.py
```

---

### Task 1: 建立可安裝的 uv package 骨架

**Files:**
- Create: `pyproject.toml`
- Create: `README.md`
- Create: `LICENSE`
- Create: `src/outlook_client/__init__.py`
- Create: `src/outlook_client/py.typed`
- Create: `tests/unit/test_package_import.py`

**Interfaces:**
- Consumes: 無。
- Produces: 可用 `uv sync` 安裝、可 import 的 `outlook_client` package。

- [ ] **Step 1: 建立 repo 並初始化 uv 專案**

```powershell
mkdir outlook-client
cd outlook-client
git init
uv init --lib --name outlook-client --python 3.11
```

- [ ] **Step 2: 將 `pyproject.toml` 設定為 src layout 與 Windows-only runtime dependency**

```toml
[project]
name = "outlook-client"
version = "0.1.0"
description = "Reusable Outlook Desktop mail search and attachment downloader"
readme = "README.md"
requires-python = ">=3.11"
dependencies = [
  "pywin32>=306; sys_platform == 'win32'",
]

[dependency-groups]
dev = [
  "mypy>=1.11",
  "pytest>=8.3",
  "pytest-cov>=5.0",
  "ruff>=0.7",
]

[build-system]
requires = ["hatchling"]
build-backend = "hatchling.build"

[tool.hatch.build.targets.wheel]
packages = ["src/outlook_client"]

[tool.pytest.ini_options]
testpaths = ["tests"]
markers = ["outlook_integration: requires Windows and a configured Outlook profile"]

[tool.ruff]
line-length = 100
target-version = "py311"

[tool.mypy]
python_version = "3.11"
strict = true
packages = ["outlook_client"]
```

- [ ] **Step 3: 寫入第一個失敗測試**

```python
def test_package_imports_without_loading_pywin32() -> None:
    import outlook_client

    assert outlook_client.__version__ == "0.1.0"
```

- [ ] **Step 4: 建立最小 `__init__.py` 並執行測試**

```python
__version__ = "0.1.0"
```

Run: `uv run pytest tests/unit/test_package_import.py -v`

Expected: PASS。

- [ ] **Step 5: 執行基礎品質檢查並提交**

```powershell
uv run ruff check .
uv run mypy src
git add pyproject.toml uv.lock README.md LICENSE src tests
git commit -m "build: initialize outlook client package"
```

---

### Task 2: 定義資料模型、查詢驗證與例外

**Files:**
- Create: `src/outlook_client/models.py`
- Create: `src/outlook_client/exceptions.py`
- Test: `tests/unit/test_models.py`

**Interfaces:**
- Consumes: Python standard library。
- Produces: `MailQuery`、`MailMessage`、`AttachmentInfo`、`ConflictPolicy` 與 package 例外 hierarchy。

- [ ] **Step 1: 寫入 query 正規化與驗證測試**

```python
from datetime import datetime

import pytest

from outlook_client.models import MailQuery


def test_query_normalizes_extensions() -> None:
    query = MailQuery(attachment_extensions=("XLSX", ".xls", ".XLSX"))
    assert query.attachment_extensions == (".xlsx", ".xls")


def test_query_rejects_reversed_time_range() -> None:
    with pytest.raises(ValueError, match="received_after"):
        MailQuery(
            received_after=datetime(2026, 9, 22),
            received_before=datetime(2026, 9, 21),
        )


def test_query_rejects_timezone_aware_datetime() -> None:
    aware = datetime.fromisoformat("2026-09-21T08:00:00+08:00")
    with pytest.raises(ValueError, match="naive local datetime"):
        MailQuery(received_after=aware)


def test_query_rejects_exact_and_contains_for_same_field() -> None:
    with pytest.raises(ValueError, match="sender"):
        MailQuery(sender="a@example.com", sender_contains="example")
```

- [ ] **Step 2: 執行測試確認失敗**

Run: `uv run pytest tests/unit/test_models.py -v`

Expected: FAIL，因 models 尚不存在。

- [ ] **Step 3: 實作資料模型的明確介面**

```python
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
    attachments: tuple[AttachmentInfo, ...]


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
```

在 `__post_init__()` 使用 `object.__setattr__()` 正規化副檔名，並驗證：時間順序、時間值皆為 naive local datetime、folder 非空、exact 與 contains 不可同時指定。

- [ ] **Step 4: 實作例外 hierarchy**

```python
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
```

- [ ] **Step 5: 跑測試與提交**

Run: `uv run pytest tests/unit/test_models.py -v`

Expected: PASS。

```powershell
git add src/outlook_client/models.py src/outlook_client/exceptions.py tests/unit/test_models.py
git commit -m "feat: define mail query and result models"
```

---

### Task 3: 實作 Outlook folder path 解析

**Files:**
- Create: `src/outlook_client/folders.py`
- Test: `tests/unit/test_folders.py`

**Interfaces:**
- Consumes: 類 Outlook Namespace 的 `GetDefaultFolder(6)` 與 folder 的 `Folders.Item(name)`。
- Produces: `split_folder_path(path: str) -> tuple[str, ...]`、`resolve_folder(namespace: Any, path: str) -> Any`。

- [ ] **Step 1: 寫入中文、英文與斜線正規化測試**

```python
import pytest

from outlook_client.exceptions import FolderNotFoundError
from outlook_client.folders import split_folder_path


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        ("收件匣/外倉/MSS", ("外倉", "MSS")),
        (r"收件匣\外倉\MSS", ("外倉", "MSS")),
        ("Inbox/Shipping Schedule", ("Shipping Schedule",)),
        ("外倉/MSS", ("外倉", "MSS")),
    ],
)
def test_split_folder_path(value: str, expected: tuple[str, ...]) -> None:
    assert split_folder_path(value) == expected


def test_rejects_empty_path_segment() -> None:
    with pytest.raises(ValueError):
        split_folder_path("收件匣//MSS")
```

- [ ] **Step 2: 寫入 fake folder resolver 測試**

建立小型 `FakeFolders`，模擬 COM collection 的 `.Item(name)`；驗證不存在的 `收件匣/不存在` 會轉成 `FolderNotFoundError`，錯誤訊息包含完整 path。

- [ ] **Step 3: 實作 resolver**

```python
OL_FOLDER_INBOX = 6


def split_folder_path(path: str) -> tuple[str, ...]:
    normalized = path.replace("\\", "/")
    parts = tuple(part.strip() for part in normalized.split("/"))
    if not parts or any(not part for part in parts):
        raise ValueError("folder path contains an empty segment")
    if parts[0].casefold() in {"inbox", "收件匣"}:
        return parts[1:]
    return parts


def resolve_folder(namespace: object, path: str) -> object:
    folder = namespace.GetDefaultFolder(OL_FOLDER_INBOX)
    try:
        for part in split_folder_path(path):
            folder = folder.Folders.Item(part)
    except Exception as exc:
        raise FolderNotFoundError(f"Outlook folder not found: {path}") from exc
    return folder
```

- [ ] **Step 4: 跑測試與提交**

Run: `uv run pytest tests/unit/test_folders.py -v`

Expected: PASS。

```powershell
git add src/outlook_client/folders.py tests/unit/test_folders.py
git commit -m "feat: resolve inbox folder paths"
```

---

### Task 4: 實作 Outlook 候選縮小與純 Python matcher

**Files:**
- Create: `src/outlook_client/filters.py`
- Test: `tests/unit/test_filters.py`

**Interfaces:**
- Consumes: `MailQuery`、`MailMessage`。
- Produces: `build_restrict_filter(query: MailQuery) -> str | None`、`message_matches(message: MailMessage, query: MailQuery) -> bool`、`attachment_matches(...) -> bool`。

- [ ] **Step 1: 寫入 Restrict filter 測試**

```python
from datetime import datetime

from outlook_client.filters import build_restrict_filter
from outlook_client.models import MailQuery


def test_builds_received_and_boolean_filter() -> None:
    query = MailQuery(
        received_after=datetime(2026, 9, 20, 8, 30),
        received_before=datetime(2026, 9, 21, 18, 0),
        has_attachment=True,
        unread_only=True,
    )
    result = build_restrict_filter(query)
    assert "[ReceivedTime] >= '09/20/2026 08:30 AM'" in result
    assert "[ReceivedTime] <= '09/21/2026 06:00 PM'" in result
    assert "[HasAttachment] = True" in result
    assert "[UnRead] = True" in result
```

- [ ] **Step 2: 寫入 sender、subject 與附件 AND 語意測試**

建立一封含 `Daily Inventory.XLSX` 的 `MailMessage`，驗證：

- `sender` 與 `sender_contains` 是不分大小寫比對。
- `subject_contains` 是不分大小寫比對。
- 同時指定附件檔名條件與副檔名時，必須由同一個附件同時符合。
- `has_attachment=False` 只接受 attachments 為空的信件。

- [ ] **Step 3: 實作 date formatter 與 Restrict builder**

```python
def format_outlook_datetime(value: datetime) -> str:
    return value.strftime("%m/%d/%Y %I:%M %p")


def build_restrict_filter(query: MailQuery) -> str | None:
    clauses: list[str] = []
    if query.received_after is not None:
        value = format_outlook_datetime(query.received_after)
        clauses.append(f"[ReceivedTime] >= '{value}'")
    if query.received_before is not None:
        value = format_outlook_datetime(query.received_before)
        clauses.append(f"[ReceivedTime] <= '{value}'")
    if query.has_attachment is not None:
        clauses.append(f"[HasAttachment] = {query.has_attachment}")
    if query.unread_only:
        clauses.append("[UnRead] = True")
    return " AND ".join(clauses) or None
```

時間字串格式需在 Task 9 的實機 smoke test 驗證公司電腦的 Outlook locale；如果台灣 Outlook 不接受此格式，僅替換這個單一 formatter，不改 public API。

- [ ] **Step 4: 實作純 Python matcher 並跑測試**

Run: `uv run pytest tests/unit/test_filters.py -v`

Expected: PASS。

- [ ] **Step 5: 提交**

```powershell
git add src/outlook_client/filters.py tests/unit/test_filters.py
git commit -m "feat: add outlook and python mail filters"
```

---

### Task 5: 建立可測試的 COM session 與信件映射

**Files:**
- Create: `src/outlook_client/protocols.py`
- Create: `src/outlook_client/com_backend.py`
- Test: `tests/unit/test_com_backend.py`

**Interfaces:**
- Consumes: lazy-loaded `pythoncom` 與 `win32com.client`。
- Produces: `ComSessionFactory` protocol、`Pywin32SessionFactory.session()` context manager、`mail_item_to_message(item, store_id) -> MailMessage`、`get_sender_email(item) -> str | None`。

- [ ] **Step 1: 定義注入介面**

```python
from contextlib import AbstractContextManager
from typing import Any, Protocol


class ComSessionFactory(Protocol):
    def session(self) -> AbstractContextManager[Any]: ...
```

- [ ] **Step 2: 寫入 COM lifecycle 測試**

使用 monkeypatch 注入 fake `pythoncom` 與 fake `win32com.client`，驗證呼叫順序為：

```text
CoInitialize
Dispatch("Outlook.Application")
GetNamespace("MAPI")
yield namespace
CoUninitialize
```

另測試 `Dispatch` 失敗時仍呼叫 `CoUninitialize()`，並丟出 `OutlookConnectionError`。

- [ ] **Step 3: 實作 lazy import 與 session context manager**

```python
class Pywin32SessionFactory:
    @contextmanager
    def session(self) -> Iterator[Any]:
        try:
            import pythoncom
            import win32com.client
        except ImportError as exc:
            raise OutlookUnavailableError(
                "pywin32 and Outlook Desktop are required on Windows"
            ) from exc

        pythoncom.CoInitialize()
        try:
            application = win32com.client.Dispatch("Outlook.Application")
            yield application.GetNamespace("MAPI")
        except OutlookError:
            raise
        except Exception as exc:
            raise OutlookConnectionError("Could not open Outlook MAPI session") from exc
        finally:
            pythoncom.CoUninitialize()
```

- [ ] **Step 4: 寫入 Exchange SMTP 解析測試**

測試三種輸入：

1. `SenderEmailType == "SMTP"`：回傳 `SenderEmailAddress`。
2. `SenderEmailType == "EX"` 且 `GetExchangeUser().PrimarySmtpAddress` 有值：回傳 SMTP address。
3. Exchange user 解析失敗：回傳原始 `SenderEmailAddress`，不讓整封信搜尋失敗。

- [ ] **Step 5: 實作 item mapper**

只映射 `Class == 43` 的 MailItem；attachment index 使用 Outlook 的 1-based index。對不存在的可選屬性用小型 `_safe_getattr()` 回退，不能吞掉必要欄位 `EntryID`、`Subject`、`ReceivedTime`。

- [ ] **Step 6: 跑測試與提交**

Run: `uv run pytest tests/unit/test_com_backend.py -v`

Expected: PASS，且 coverage 包含正常與例外 cleanup path。

```powershell
git add src/outlook_client/protocols.py src/outlook_client/com_backend.py tests/unit/test_com_backend.py
git commit -m "feat: isolate outlook com sessions"
```

---

### Task 6: 實作 `search()` 與 `find_latest()`

**Files:**
- Create: `src/outlook_client/client.py`
- Test: `tests/unit/test_client.py`

**Interfaces:**
- Consumes: `ComSessionFactory`、folder resolver、Restrict builder、COM mapper、Python matcher。
- Produces: `OutlookClient.search(query, limit=None) -> list[MailMessage]`、`OutlookClient.find_latest(query) -> MailMessage | None`。

- [ ] **Step 1: 寫入 fake COM collection 測試**

測試必須驗證：

- 有 filter 時呼叫 `Items.Restrict(filter_string)`。
- 無 Outlook-side 條件時不呼叫 `Restrict()`。
- 呼叫 `Items.Sort("[ReceivedTime]", True)`，結果由新到舊。
- 非 MailItem (`Class != 43`) 被忽略。
- Python matcher 過濾 sender、subject 與附件。
- `limit=2` 時只回傳兩封，且 `limit=0` 或負數丟出 `ValueError`。

- [ ] **Step 2: 實作 constructor 與 search**

```python
class OutlookClient:
    def __init__(self, session_factory: ComSessionFactory | None = None) -> None:
        self._session_factory = session_factory or Pywin32SessionFactory()

    def search(self, query: MailQuery, *, limit: int | None = None) -> list[MailMessage]:
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
                if getattr(item, "Class", None) != 43:
                    continue
                message = mail_item_to_message(item, folder.StoreID)
                if message_matches(message, query):
                    results.append(message)
                    if limit is not None and len(results) >= limit:
                        break
        return results
```

- [ ] **Step 3: 實作 latest semantics**

```python
def find_latest(self, query: MailQuery) -> MailMessage | None:
    messages = self.search(query, limit=1)
    return messages[0] if messages else None
```

- [ ] **Step 4: 跑測試與提交**

Run: `uv run pytest tests/unit/test_client.py -v`

Expected: PASS。

```powershell
git add src/outlook_client/client.py tests/unit/test_client.py
git commit -m "feat: search outlook mail and find latest"
```

---

### Task 7: 實作附件篩選與同名檔策略

**Files:**
- Create: `src/outlook_client/attachments.py`
- Test: `tests/unit/test_attachments.py`

**Interfaces:**
- Consumes: `AttachmentInfo`、`ConflictPolicy`、`Path`。
- Produces: `select_attachments(attachments, *, filename=None, filename_contains=None, extensions=()) -> tuple[AttachmentInfo, ...]`、`safe_attachment_name(filename: str) -> str`、`resolve_destination(...) -> Path | None`。

- [ ] **Step 1: 寫入附件選取測試**

```python
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
```

- [ ] **Step 2: 寫入四種 conflict policy 測試**

使用 pytest `tmp_path` 驗證：

- `overwrite` 回傳原路徑。
- `skip` 回傳 `None`。
- `error` 丟出 `AttachmentConflictError`。
- `rename` 依序產生 `庫存表_1.xlsx`、`庫存表_2.xlsx`，保留複合 suffix 前的 stem。

- [ ] **Step 3: 防止附件檔名 path traversal**

Outlook 附件名稱即使包含 `../`、`..\` 或絕對路徑，也只能使用 `PureWindowsPath(filename).name` 的安全 basename；另拒絕 `""`、`"."` 與 `".."`。加入測試確認輸出仍位於 `output_dir.resolve()` 之下，並讓測試在非 Windows CI 也有相同行為。

- [ ] **Step 4: 實作純函式並跑測試**

Run: `uv run pytest tests/unit/test_attachments.py -v`

Expected: PASS。

- [ ] **Step 5: 提交**

```powershell
git add src/outlook_client/attachments.py tests/unit/test_attachments.py
git commit -m "feat: filter attachments and resolve conflicts"
```

---

### Task 8: 實作附件下載與 `download_latest()`

**Files:**
- Modify: `src/outlook_client/client.py`
- Modify: `tests/unit/test_client.py`

**Interfaces:**
- Consumes: `MailMessage` locator、附件 selector、destination resolver、Outlook `Attachment.SaveAsFile(path)`。
- Produces: `download_attachments(...) -> list[Path]`、`download_latest(...) -> list[Path]`。

- [ ] **Step 1: 寫入成功下載測試**

```python
def test_download_attachments_returns_saved_paths(tmp_path: Path) -> None:
    client = OutlookClient(session_factory=fake_session_factory)
    paths = client.download_attachments(
        message,
        output_dir=tmp_path,
        extensions=(".xlsx",),
        conflict="error",
    )
    assert paths == [tmp_path / "每日庫存.xlsx"]
    fake_attachment.SaveAsFile.assert_called_once_with(str(paths[0].resolve()))
```

- [ ] **Step 2: 寫入錯誤轉譯測試**

驗證：

- `GetItemFromID()` 失敗會丟出 `MailAccessError`。
- 沒有符合附件會丟出 `AttachmentNotFoundError`。
- `SaveAsFile()` 失敗會丟出 `AttachmentSaveError`，訊息包含 filename 與 destination。
- `skip` 導致全部附件跳過時回傳空 list，不視為找不到附件。

- [ ] **Step 3: 實作 download method signature**

```python
def download_attachments(
    self,
    message: MailMessage,
    output_dir: str | os.PathLike[str],
    *,
    extensions: tuple[str, ...] = (),
    filename: str | None = None,
    filename_contains: str | None = None,
    conflict: ConflictPolicy = "error",
) -> list[Path]: ...
```

執行順序固定為：建立 output directory → 選出 metadata → 重新取得 mail item → 以 metadata 的 1-based index 取得附件 → 核對目前 `FileName` 仍與 metadata 相同 → 決定 destination → `SaveAsFile(str(destination.resolve()))` → 收集成功路徑。若信件附件在 search 與 download 之間改變，丟出 `MailAccessError`，避免下載錯誤附件。

- [ ] **Step 4: 實作 high-level API**

```python
def download_latest(
    self,
    query: MailQuery,
    output_dir: str | os.PathLike[str],
    *,
    conflict: ConflictPolicy = "error",
) -> list[Path]:
    message = self.find_latest(query)
    if message is None:
        raise MailNotFoundError("No Outlook message matched the query")
    return self.download_attachments(
        message,
        output_dir,
        extensions=query.attachment_extensions,
        filename=query.attachment_name,
        filename_contains=query.attachment_name_contains,
        conflict=conflict,
    )
```

- [ ] **Step 5: 跑測試與提交**

Run: `uv run pytest tests/unit/test_client.py tests/unit/test_attachments.py -v`

Expected: PASS。

```powershell
git add src/outlook_client/client.py tests/unit/test_client.py
git commit -m "feat: download matching outlook attachments"
```

---

### Task 9: 穩定 public exports、logging 與使用文件

**Files:**
- Modify: `src/outlook_client/__init__.py`
- Modify: `README.md`
- Create: `tests/unit/test_public_api.py`

**Interfaces:**
- Consumes: 前面所有 public types。
- Produces: 穩定的 `from outlook_client import ...` 介面與跨專案安裝說明。

- [ ] **Step 1: 寫入 public API 測試**

```python
from outlook_client import (
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
```

- [ ] **Step 2: 僅匯出穩定介面**

`__init__.py` 匯出：

```python
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
```

不要匯出 `Pywin32SessionFactory`、folder resolver、filter builder 或 COM constants。

- [ ] **Step 3: 補 logging**

各 module 使用：

```python
import logging

logger = logging.getLogger(__name__)
```

建議級別：query 摘要與命中數用 `DEBUG`；附件成功下載用 `INFO`；skip 用 `INFO`；sender SMTP fallback 用 `DEBUG`；不得記錄信件本文。

- [ ] **Step 4: README 寫入完整範例**

README 至少包含：

- Windows／Classic Outlook 前置條件。
- `uv add` 的本機 editable、Git URL 與正式 registry 三種方式。
- `search()`、`find_latest()`、`download_attachments()`、`download_latest()` 範例。
- 四種 conflict policy 行為表。
- 例外處理範例，特別是 `MailNotFoundError` 對應「尚未收到今日資料」。
- package 與業務專案的責任邊界。
- worker thread 的 COM lifecycle 已由 package 處理，呼叫端不需自行 `CoInitialize()`。

- [ ] **Step 5: 跑測試與提交**

Run: `uv run pytest tests/unit -v`

Expected: PASS。

```powershell
git add src/outlook_client README.md tests/unit/test_public_api.py
git commit -m "docs: publish stable outlook client api"
```

---

### Task 10: 在真實 Windows Outlook 執行 opt-in smoke test

**Files:**
- Create: `tests/integration/test_outlook_smoke.py`
- Modify: `README.md`

**Interfaces:**
- Consumes: 已設定 Outlook profile 的 Windows 11 電腦。
- Produces: 可確認 COM、folder、Restrict、中文路徑及實際附件下載的 opt-in 測試。

- [ ] **Step 1: 建立不會預設執行的 smoke test**

```python
import os
from datetime import datetime, timedelta

import pytest

from outlook_client import MailQuery, OutlookClient


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
```

- [ ] **Step 2: 在公司 Windows 11 電腦跑搜尋 smoke test**

```powershell
$env:RUN_OUTLOOK_INTEGRATION = "1"
uv run pytest tests/integration/test_outlook_smoke.py -v -s
```

Expected: Outlook profile 已設定時測試 PASS，沒有 recent mail 時回傳空 list 仍 PASS。

- [ ] **Step 3: 驗證台灣 locale 的 ReceivedTime Restrict**

以 Outlook UI 已知的一封 24 小時內信件交叉確認 search 能找到；再把 `received_after` 移到該信件之後，確認找不到。若第一項失敗，修改 `format_outlook_datetime()`，並把公司電腦實際接受的格式固化為 unit test。

- [ ] **Step 4: 驗證 opt-in 附件下載**

使用專用測試信件與暫存資料夾；先 `find_latest()` 顯示 subject、sender、attachment filenames，再由測試者明確設定 `RUN_OUTLOOK_DOWNLOAD_INTEGRATION=1` 才執行 `download_attachments()`。測試不得刪除、搬移或修改信件。

- [ ] **Step 5: 提交**

```powershell
git add tests/integration/test_outlook_smoke.py README.md
git commit -m "test: add opt-in outlook smoke coverage"
```

---

### Task 11: 完整驗證、建立 wheel 與跨專案試裝

**Files:**
- Modify: `README.md` only if verification exposes inaccurate instructions.

**Interfaces:**
- Consumes: 完成的 package。
- Produces: 可由其他 uv project 安裝的 wheel 與驗證紀錄。

- [ ] **Step 1: 執行所有非實機測試**

```powershell
uv run pytest -m "not outlook_integration" --cov=outlook_client --cov-report=term-missing
```

Expected: 全部 PASS；核心純函式與 client orchestration coverage 不低於 90%。

- [ ] **Step 2: 執行靜態檢查**

```powershell
uv run ruff format --check .
uv run ruff check .
uv run mypy src
```

Expected: 全部 exit code 0。

- [ ] **Step 3: 建立並檢查 wheel**

```powershell
uv build
uvx twine check dist\*
```

Expected: 產生 source distribution 與 wheel，metadata check PASS。

- [ ] **Step 4: 由臨時 uv project 安裝 wheel**

```powershell
mkdir D:\Temp\outlook-client-consumer
cd D:\Temp\outlook-client-consumer
uv init --python 3.11
uv add D:\Code2\outlook-client\dist\outlook_client-0.1.0-py3-none-any.whl
uv run python -c "from outlook_client import OutlookClient, MailQuery; print(MailQuery())"
```

Expected: consumer project 不需要調整 `PYTHONPATH` 即可 import。

- [ ] **Step 5: 選一個真實專案做最小整合**

優先選外倉庫存下載專案，只替換「Outlook 找信與下載附件」段落；業務邏輯、檔名改名規則、資料處理與排程仍留在原專案。至少連續手動執行兩次，確認第二次的 conflict policy 符合預期。

- [ ] **Step 6: 建立 v0.1.0 tag**

```powershell
git status --short
git tag -a v0.1.0 -m "outlook client v0.1.0"
git log --oneline --decorate -5
```

Expected: worktree 乾淨，tag 指向所有驗證通過的 commit。

---

### Task 12: 支援 `MailQuery.recursive` 遞迴子資料夾搜尋

> v1 完成（Task 1–11）後追加的能力。背景：呼叫端搜尋範圍設在「收件匣」，但信件實際會落在哪個子資料夾不確定、且結構會變動，無法用固定資料夾清單在呼叫端逐一搜尋解決。由於呼叫端拿到的只有純 Python 的 `MailMessage`，完全接觸不到 COM object，遞迴走訪資料夾樹的能力只能由這個 package 提供，不能留給呼叫端自行用 `pywin32` 兜。

**Files:**
- Modify: `src/email_downloader/models.py`
- Modify: `src/email_downloader/folders.py`
- Modify: `src/email_downloader/client.py`
- Modify: `tests/units/test_folders.py`
- Modify: `tests/units/test_client.py`
- Modify: `README.md`
- Modify: `docs/progress.md`

**Interfaces:**
- Consumes: 既有 `resolve_folder()`、COM 資料夾的 `Folders` collection。
- Produces: `MailQuery.recursive: bool = False`；`folders.iter_folder_tree(folder: Any) -> Iterator[Any]`；`search()`／`find_latest()`／`download_latest()` 在 `recursive=True` 時搜尋 `query.folder` 底下所有層級的子資料夾（不限深度）。

**設計要點：**

- `recursive` 預設 `False`，維持現有單一資料夾行為完全不變，屬於選擇性加入的能力。
- `query.folder` 是遞迴搜尋的**起點**；`iter_folder_tree()` 以 DFS 走訪該資料夾自己與底下所有層級的子資料夾，不限深度。
- `recursive=False`：邏輯不變，維持「蒐集到 `limit` 筆即提早中斷」的效能優化（單一資料夾已用 `Items.Sort()` 由新到舊排序，提早中斷是安全的）。
- `recursive=True`：**不可提早中斷**——必須先搜完 `query.folder` 底下所有子資料夾、蒐集全部符合條件的信件，再依 `received_time` 做一次全域排序，最後才依 `limit` 截斷。原因是在還沒搜完所有資料夾前，無法確定哪一封信是全域最新的。
- 錯誤處理：遞迴過程中若某個子資料夾本身存取失敗（例如權限問題），比照 Task 6 review「跳過讀不到的單一信件」的既有原則——跳過該子資料夾（記 log）、繼續搜其他資料夾，不讓整個 `search()` 失敗。只有 `query.folder` 這個根路徑本身不存在時，才維持原本的 `FolderNotFoundError`。
- `find_latest()`／`download_latest()` 不需要另外改介面：它們本來就是直接把 `query` 傳下去，`recursive` 會自動一路帶過去。

- [ ] **Step 1: 寫入 `iter_folder_tree()` 的多層 fake `.Folders` 測試**

驗證走訪順序涵蓋根資料夾自己與所有層級的子資料夾（例如根下有兩個子資料夾，其中一個底下還有孫層資料夾）。

- [ ] **Step 2: 實作 `iter_folder_tree()`**

```python
def iter_folder_tree(folder: Any) -> Iterator[Any]:
    yield folder
    for sub in folder.Folders:
        yield from iter_folder_tree(sub)
```

- [ ] **Step 3: 寫入 `search(recursive=True)` 的多資料夾 fake 測試**

至少涵蓋：

- 兩個子資料夾各自命中信件時，結果依 `received_time` 全域排序、正確合併。
- `limit` 在全域排序「之後」才截斷，不會漏掉排序較後被搜到、但實際時間較新的信件。
- 某個子資料夾的 `.Items` 存取拋例外時，該資料夾被跳過，其餘資料夾結果不受影響。
- `recursive=False`（預設）的既有行為與提早中斷優化不受影響（迴歸測試）。

- [ ] **Step 4: 實作 `client.py` 的 `recursive` 分支邏輯**

- [ ] **Step 5: 跑測試、更新 README 與 `docs/progress.md`、提交**

```powershell
uv run pytest tests/units -v
uv run ruff check .
uv run mypy src
git add src/email_downloader/models.py src/email_downloader/folders.py src/email_downloader/client.py tests/units/test_folders.py tests/units/test_client.py README.md docs/progress.md
git commit -m "feat: support recursive subfolder search"
```

---

## Definition of Done

- [ ] 其他專案可用 `uv add --editable D:\Code2\outlook-client` 或 Git dependency 安裝。
- [ ] 可依 folder、sender、subject、received time、attachment filename、extension 與 unread 搜尋。
- [ ] `search()` 回傳純 Python models，沒有 COM object 外洩。
- [ ] `find_latest()` 與 `download_latest()` 行為符合設計基準。
- [ ] 四種同名檔策略都有 unit tests。
- [ ] 自訂例外完整包住 pywin32／COM 錯誤。
- [ ] worker thread 的 COM initialization／cleanup 有測試。
- [ ] 台灣公司電腦的 Classic Outlook smoke test 通過。
- [ ] `pytest`、`ruff`、`mypy`、wheel build 全部通過。
- [ ] README 足以讓另一個 project 在不閱讀 package internals 的情況下使用。

## 第二階段候選項目

只有在 v0.1.0 實際被兩個以上專案使用後再評估：

1. YAML config 與 `download_from_config()`。
2. `body_contains`、recipient、CC、categories、importance。
3. shared mailbox／指定 store。
4. hash-based 去重與下載 manifest。
5. CLI。
6. Graph 或 IMAP backend；保留同一組 `MailQuery` 與 models，但不強迫不同 backend 共用 COM internal protocol。

## 官方技術參考

- Microsoft Learn — Items.Restrict: https://learn.microsoft.com/en-us/office/vba/api/outlook.items.restrict
- Microsoft Learn — NameSpace.GetItemFromID: https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace.getitemfromid
- Microsoft Learn — Attachment.SaveAsFile: https://learn.microsoft.com/en-us/office/vba/api/outlook.attachment.saveasfile
- Microsoft Learn — MailItem.SenderEmailAddress: https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.senderemailaddress
