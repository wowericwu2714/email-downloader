# Progress

對照 `docs/outlook-client-PLAN.md` 的 Task 清單追蹤進度。

- [x] Task 1: 建立可安裝的 uv package 骨架（commit `503574c`）
- [x] Task 2: 定義資料模型、查詢驗證與例外（commit `f171f5b`）
- [ ] Task 3: 實作 Outlook folder path 解析
- [ ] Task 4: 實作 Outlook 候選縮小與純 Python matcher
- [ ] Task 5: 建立可測試的 COM session 與信件映射
- [ ] Task 6: 實作 `search()` 與 `find_latest()`
- [ ] Task 7: 實作附件篩選與同名檔策略
- [ ] Task 8: 實作附件下載與 `download_latest()`
- [ ] Task 9: 穩定 public exports、logging 與使用文件
- [ ] Task 10: 在真實 Windows Outlook 執行 opt-in smoke test
- [ ] Task 11: 完整驗證、建立 wheel 與跨專案試裝

## 待確認事項

- Package import 名稱目前是 `email_downloader`，計畫原文是 `outlook_client`，尚未統一。

## 最後驗證狀態（commit `f171f5b`）

- `uv run pytest tests/units -v`：5 passed
- `uv run ruff check .`：All checks passed
- `uv run mypy src`：Success
