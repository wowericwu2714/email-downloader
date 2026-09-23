# Progress

對照 `docs/email-downloader-PLAN.md` 的 Task 清單追蹤進度。

- [x] Task 1: 建立可安裝的 uv package 骨架（commit `503574c`）
- [x] Task 2: 定義資料模型、查詢驗證與例外（commit `f171f5b`）
- [x] Task 3: 實作 Outlook folder path 解析（commit `4535abd`，main）
- [x] Task 4: 實作 Outlook 候選縮小與純 Python matcher（commit `26217ec`，main）
- [ ] Task 5: 建立可測試的 COM session 與信件映射（進行中，分支 `feature/task5_com-backend`）
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

## Task 5 進度（進行中，未 commit）

分支 `feature/task5_com-backend`，目前已 staged 尚未 commit：

- ✅ `src/email_downloader/com_backend.py`：`Pywin32SessionFactory`（COM session 生命週期：lazy import → `CoInitialize` → `Dispatch` → `GetNamespace` → yield → `CoUninitialize`，含 Dispatch 失敗時的例外轉譯）、`get_sender_email()`（Exchange `EX` 位址優先解析 `PrimarySmtpAddress`，失敗時安全退回原始位址）。對應 Review Focus 第 1、2 項，測試已覆蓋。
- ⏳ `mail_item_to_message()` 尚未實作：`tests/units/test_com_backend.py` 已經 import 這個函式，導致目前 `tests/units` 整個目錄都無法收集（`ImportError`）。測試檔尾端的 `FakeAttachment`／`FakeAttachments` 是為它準備的假物件，還沒有任何測試使用。
- `pyproject.toml` 新增 dev dependency `types-pywin32`，供 mypy 使用。

## 最後驗證狀態（2026-09-23，Task 5 進行中，紅燈）

- `uv run pytest tests/units -v`：collection error（`mail_item_to_message` 尚未定義於 `com_backend.py`），24 個既有測試因此全部無法收集
- `uv run ruff check .`：1 個 F401（`mail_item_to_message` imported but unused）
- `uv run mypy src`：尚未執行（pytest 先擋下）
