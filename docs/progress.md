# Progress

對照 `docs/email-downloader-PLAN.md` 的 Task 清單追蹤進度。

- [x] Task 1: 建立可安裝的 uv package 骨架（commit `503574c`）
- [x] Task 2: 定義資料模型、查詢驗證與例外（commit `f171f5b`）
- [x] Task 3: 實作 Outlook folder path 解析（commit `4535abd`，main）
- [x] Task 4: 實作 Outlook 候選縮小與純 Python matcher（commit `26217ec`，main）
- [x] Task 5: 建立可測試的 COM session 與信件映射（commit `d4ef53c`，`feature/task5_com-backend`；已合併至 `origin/main` `afc618c`）
- [ ] Task 6: 實作 `search()` 與 `find_latest()`
- [ ] Task 7: 實作附件篩選與同名檔策略
- [ ] Task 8: 實作附件下載與 `download_latest()`
- [ ] Task 9: 穩定 public exports、logging 與使用文件
- [ ] Task 10: 在真實 Windows Outlook 執行 opt-in smoke test
- [ ] Task 11: 完整驗證、建立 wheel 與跨專案試裝

## 待確認事項

- ~~Package import 名稱~~：已確定維持 `email_downloader`；plan 文件已改名為 `docs/email-downloader-PLAN.md`，但本文內部分段落仍沿用舊的 `outlook_client` 措辭，屬文件小瑕疵，不影響開發。

## Task 4 code review 待辦事項（Task 9 前需確認）

- **`format_outlook_datetime` 的地區設定風險**（`filters.py`）：目前用固定美式格式（`MM/DD/YYYY hh:mm AM/PM`）餵給 Outlook `Items.Restrict()`。PLAN 文件本身已預告此風險，需在 Task 9 實機 smoke test 時，用公司電腦（繁中 Windows）驗證 Outlook 能否正確解析；若不行，只需替換 `format_outlook_datetime()` 這一個函式，不影響 public API。
- **`format_outlook_datetime` 只精確到分鐘**（`filters.py`）：`Items.Restrict` 用的這種日期字串格式本身不支援秒級精度，屬於此方案的固有限制。目前業務情境（每日庫存信）不需要秒級篩選，暫不處理。

## Task 5 完成內容（`src/email_downloader/com_backend.py`）

- `Pywin32SessionFactory`：COM session 生命週期（lazy import → `CoInitialize` → `Dispatch` → `GetNamespace` → yield → `CoUninitialize`，含 Dispatch 失敗時的例外轉譯）。
- `get_sender_email()`：Exchange `EX` 位址優先解析 `PrimarySmtpAddress`，失敗時安全退回原始位址。
- `mail_item_to_message()`：COM `MailItem` → 純 Python `MailMessage`。`EntryID`／`Subject`／`ReceivedTime` 直接存取（缺漏時讓例外自然噴出，不吞掉），其餘可選欄位透過新增的 `_safe_getattr()` 安全讀取；附件用 Outlook 1-based index，副檔名經 `PureWindowsPath(...).suffix.casefold()` 正規化。
- `pyproject.toml` 新增 dev dependency `types-pywin32`，供 mypy 使用。
- 對應 Review Focus 第 1、2 項，測試已全數覆蓋。

## 待辦：`protocols.py` 尚未被使用

`ComSessionFactory` Protocol 已定義但目前沒有任何程式碼引用，要等 Task 6 `OutlookClient.__init__(session_factory: ComSessionFactory | None = None)` 才會用上，屬預期中的狀態。

## 最後驗證狀態（2026-09-23，Task 5 完成）

- `uv run pytest tests/units -v`：33 passed
- `uv run ruff check .`：All checks passed
- `uv run mypy src`：Success（7 source files）
- 本地 `main` 分支落後 `origin/main` 一個 commit（`afc618c`，與 `feature/task5_com-backend` 內容相同），尚未 `git pull` 同步。
