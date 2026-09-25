# email-downloader

`email-downloader`（import 名稱：`email_downloader`）是一個可重複使用的 Windows Outlook Desktop 郵件搜尋與附件下載 Python package，設計給 ICS、外倉庫存、Shipping Schedule 等內部專案直接引入。

## 責任邊界

- 這個 package **只負責「如何找信、如何下載附件」**：解析 Outlook 資料夾路徑、用 `Items.Restrict` 縮小候選信件、純 Python 比對 sender／subject／附件檔名、下載附件並處理同名檔衝突。
- **業務條件留在呼叫端**：哪個資料夾對應哪個業務流程、找不到今日資料時要不要沿用舊資料、下載後檔名要不要改名、下載紀錄要不要寫資料庫，這些都由呼叫端的專案自行決定，本 package 不涉入。
- 對外只回傳純 Python 的 `MailQuery`、`MailMessage`、`AttachmentInfo`、`Path` 與例外物件，不會回傳 `win32com` COM object；下載時會在同一個 COM session 內用 `EntryID`／`StoreID` 重新取得信件。

## 前置條件

- Windows 11。
- Classic Outlook Desktop（非 New Outlook）已安裝並登入好一個 profile。
- Python 3.11+。
- 在非 Windows 環境（例如 CI、macOS 開發機）仍可以 `import email_downloader` 並執行不涉及 COM 的單元測試——`pywin32`／`pythoncom` 是 lazy import，只有實際呼叫 `search()`／`download_attachments()` 等會觸碰 COM 的方法時才需要 Windows。

## 安裝

使用 `uv` 時依情境挑一種：

```bash
# 本機 editable（開發、或本機還沒發到任何 registry 前）
uv add --editable D:\Code2\email-downloader

# Git dependency（跨專案最常用的方式）
uv add git+https://github.com/wowericwu2714/email-downloader.git

# 已發佈到正式 registry 後
uv add email-downloader
```

## 快速開始

### 高階 API：一次找信並下載附件

```python
from datetime import datetime, timedelta
from email_downloader import MailQuery, OutlookClient

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

### 低階 API：先檢查信件內容，再決定要不要下載

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

## API 一覽

### `OutlookClient()`

不帶參數即可建立；worker thread（Prefect、APScheduler 等）的 COM lifecycle（`CoInitialize`／`CoUninitialize`）已由 package 內部處理，呼叫端**不需要**自行初始化 COM。

### `search(query: MailQuery, *, limit: int | None = None) -> list[MailMessage]`

依 `MailQuery` 搜尋，由新到舊排序回傳。可搜尋條件涵蓋 folder、sender（exact／contains）、subject（exact／contains）、received time 區間、附件檔名（exact／contains）、附件副檔名與 unread 狀態。找不到符合的信件時回傳空 list。

```python
messages = client.search(
    MailQuery(folder="收件匣/外倉/MSS", sender_contains="mss"),
    limit=10,
)
```

### `find_latest(query: MailQuery) -> MailMessage | None`

回傳最新一封符合條件的信件；找不到時回傳 `None`（不會拋例外）。

### `download_attachments(message, output_dir, *, extensions=(), filename=None, filename_contains=None, conflict="error") -> list[Path]`

針對一封已經取得的 `MailMessage`（來自 `search()` 或 `find_latest()`）下載符合條件的附件，回傳成功存檔的路徑清單。

### `download_latest(query, output_dir, *, conflict="error") -> list[Path]`

`find_latest()` + `download_attachments()` 的組合：找不到信件時拋出 `MailNotFoundError`；找到信件但沒有符合的附件時拋出 `AttachmentNotFoundError`。附件篩選條件直接沿用 `query` 裡的 `attachment_name`／`attachment_name_contains`／`attachment_extensions`。

## 同名檔衝突策略（`conflict` 參數）

| policy | 行為 |
| --- | --- |
| `"overwrite"` | 直接覆蓋既有檔案，回傳原路徑。 |
| `"skip"` | 該附件不下載，不計入回傳結果，也不算失敗。 |
| `"rename"` | 依序產生 `檔名_1.xlsx`、`檔名_2.xlsx` ……，保留原始副檔名（含 `.tar.gz` 這種複合副檔名）。 |
| `"error"`（預設） | 拋出 `AttachmentConflictError`，不下載任何後續衝突的附件。 |

## 例外處理

所有 package 自訂例外都繼承自 `OutlookError`，不會外洩原始的 `pywin32`／COM 例外：

```python
from email_downloader import (
    MailNotFoundError,
    AttachmentNotFoundError,
    OutlookClient,
    MailQuery,
)

client = OutlookClient()

try:
    files = client.download_latest(
        MailQuery(folder="收件匣/外倉/MSS", subject_contains="每日庫存"),
        output_dir=r"D:\data\warehouse",
    )
except MailNotFoundError:
    # 例如：今天還沒收到外倉庫存表，沿用最近一次的結存資料
    files = []
except AttachmentNotFoundError:
    # 信件收到了，但沒有符合檔名／副檔名條件的附件
    files = []
```

其餘常見例外：

| 例外 | 何時發生 |
| --- | --- |
| `OutlookUnavailableError` | 目前環境沒有安裝 `pywin32` 或不是 Windows。 |
| `OutlookConnectionError` | 無法開啟 Outlook COM session（例如 Outlook 沒有啟動、profile 未設定）。 |
| `FolderNotFoundError` | `MailQuery.folder` 指定的路徑在 Outlook 裡找不到。 |
| `MailAccessError` | 信件在 `search()` 之後被移動、刪除，或附件內容已改變，導致 `download_attachments()` 無法安全下載。 |
| `AttachmentConflictError` | `conflict="error"` 且目的檔已存在。 |
| `AttachmentSaveError` | Outlook 的 `SaveAsFile()` 失敗；訊息會列出失敗的檔名與已經成功存下的路徑（best-effort：其他附件仍會照常嘗試存檔）。 |

## Logging

package 內部只呼叫 `logging.getLogger(__name__)`，不會呼叫 `logging.basicConfig()`，也不會記錄信件本文；呼叫端專案自行決定要不要開啟、輸出到哪裡。目前使用的層級：

- `DEBUG`：查詢摘要、Restrict filter 內容、搜尋命中數、Exchange sender SMTP fallback。
- `INFO`：附件成功下載、附件因同名檔衝突被 skip。

## v1 範圍之外

不支援：遞迴搜尋所有子資料夾、寫信／回信／搬移／刪除／標記已讀、本文搜尋、收件者／CC／分類／重要性條件、YAML／JSON 設定驅動下載、CLI、Microsoft Graph／IMAP／macOS backend、跨 mailbox 的頂層 store 選擇、下載紀錄資料庫與排程功能。
