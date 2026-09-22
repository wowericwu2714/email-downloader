# Progress

對照 `docs/outlook-client-PLAN.md` 的 Task 清單追蹤進度。

- [x] Task 1: 建立可安裝的 uv package 骨架（commit `503574c`）
- [x] Task 2: 定義資料模型、查詢驗證與例外（commit `f171f5b`）
- [x] Task 3: 實作 Outlook folder path 解析（commit `4535abd`）
- [x] Task 4: 實作 Outlook 候選縮小與純 Python matcher（實作完成，code review 已修正並複核通過，尚未 commit）
- [ ] Task 5: 建立可測試的 COM session 與信件映射
- [ ] Task 6: 實作 `search()` 與 `find_latest()`
- [ ] Task 7: 實作附件篩選與同名檔策略
- [ ] Task 8: 實作附件下載與 `download_latest()`
- [ ] Task 9: 穩定 public exports、logging 與使用文件
- [ ] Task 10: 在真實 Windows Outlook 執行 opt-in smoke test
- [ ] Task 11: 完整驗證、建立 wheel 與跨專案試裝

## 待確認事項

- Package import 名稱目前是 `email_downloader`，計畫原文是 `outlook_client`，尚未統一。

## Task 4 code review 待辦事項（Task 9 前需確認）

- **`format_outlook_datetime` 的地區設定風險**（`filters.py`）：目前用固定美式格式（`MM/DD/YYYY hh:mm AM/PM`）餵給 Outlook `Items.Restrict()`。PLAN 文件本身已預告此風險（Task 4 Step 3 註記），需在 Task 9 實機 smoke test 時，用公司電腦（繁中 Windows）驗證 Outlook 能否正確解析；若不行，只需替換 `format_outlook_datetime()` 這一個函式，不影響 public API。
- **`format_outlook_datetime` 只精確到分鐘**（`filters.py`）：`Items.Restrict` 用的這種日期字串格式本身不支援秒級精度（需改用 DASL 語法才能做到），屬於此方案的固有限制。目前業務情境（每日庫存信）不需要秒級篩選，暫不處理；若未來需要秒級精度，需評估改用 DASL 查詢語法。

## 最後驗證狀態（Task 4 完成，尚未 commit）

- `uv run pytest tests/units -v`：24 passed
- `uv run ruff check .`：All checks passed
- `uv run mypy src`：Success
