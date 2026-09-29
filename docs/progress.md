# Progress

對照 `docs/email-downloader-PLAN.md` 的 Task 清單追蹤進度。

- [x] Task 1: 建立可安裝的 uv package 骨架（commit `503574c`）
- [x] Task 2: 定義資料模型、查詢驗證與例外（commit `f171f5b`）
- [x] Task 3: 實作 Outlook folder path 解析（commit `4535abd`，main）
- [x] Task 4: 實作 Outlook 候選縮小與純 Python matcher（commit `26217ec`，main）
- [x] Task 5: 建立可測試的 COM session 與信件映射（`com_backend.py`，已合併至 main）
- [x] Task 6: 實作 `search()` 與 `find_latest()`（`client.py`，已合併至 main）
- [x] Task 7: 實作附件篩選與同名檔策略（`attachments.py`，已合併至 main）
- [x] Task 8: 實作附件下載與 `download_latest()`（`client.py` 的 `download_attachments`／`download_latest`，已合併至 main）
- [x] Task 9: 穩定 public exports、logging 與使用文件（分支 `feature/task9_public-api`，commit `3450914`）
- [x] Task 12: 支援遞迴搜尋子資料夾（commit `bbb01d0`，PR #2 已合併至 main `696cf9f`）
- [x] Task 10: 在真實 Windows Outlook 執行 opt-in smoke test（commit `d4cd806`，分支 `feature/task10_smoke-test`，尚未合併至 main）
- [ ] Task 11: 完整驗證、建立 wheel 與跨專案試裝（分支 `feature/task11_build-test`，Step 1-4 完成，Step 5/6 待決定，尚未 commit）

> Task 12 是後來才追加進 `docs/email-downloader-PLAN.md` 的新任務（原計畫的 Task 10/11 之前），所以編號沒有照順序放，這裡維持文件裡的實際 Task 編號。

> 注意：這份文件在 2026-09-24 的一次合併衝突處理（commit `ee71a4f`）中被意外還原成舊版本（連同 `com_backend.py` 的內容重複問題一起發生），2026-09-25 已依對話紀錄重新整理回正確狀態。以後解衝突時這份檔案也要仔細檢查，不能整份被舊版蓋掉。

## 待確認事項

- ~~Package import 名稱~~：已確定維持 `email_downloader`；plan 文件已改名為 `docs/email-downloader-PLAN.md`，但本文內部分段落仍沿用舊的 `outlook_client` 措辭，屬文件小瑕疵，不影響開發。

## Task 4 code review 待辦事項（Task 9/10 真機測試前需確認）

- **`format_outlook_datetime` 的地區設定風險**（`filters.py`）：目前用固定美式格式（`MM/DD/YYYY hh:mm AM/PM`）餵給 Outlook `Items.Restrict()`。PLAN 文件本身已預告此風險，需在 Task 9/10 實機 smoke test 時，用公司電腦（繁中 Windows）驗證 Outlook 能否正確解析；若不行，只需替換 `format_outlook_datetime()` 這一個函式，不影響 public API。
- **`format_outlook_datetime` 只精確到分鐘**（`filters.py`）：`Items.Restrict` 用的這種日期字串格式本身不支援秒級精度，屬於此方案的固有限制。目前業務情境（每日庫存信）不需要秒級篩選，暫不處理。

## Task 5 完成內容（`src/email_downloader/com_backend.py`）

- `Pywin32SessionFactory`：COM session 生命週期（lazy import → `CoInitialize` → `Dispatch` → `GetNamespace` → yield → `CoUninitialize`）。`session()` 的 try/except 範圍後來（Task 6 review）收窄成只包住 `Dispatch`／`GetNamespace`，避免 `with` 區塊內（`client.py` 的邏輯）丟出的例外被誤標成連線失敗。
- `get_sender_email()`：Exchange `EX` 位址優先解析 `PrimarySmtpAddress`，失敗時安全退回原始位址。
- `mail_item_to_message()`：COM `MailItem` → 純 Python `MailMessage`。`EntryID`／`Subject`／`ReceivedTime` 直接存取（缺漏時讓例外自然噴出，不吞掉），其餘可選欄位透過 `_safe_getattr()` 安全讀取；附件用 Outlook 1-based index，副檔名經 `PureWindowsPath(...).suffix.casefold()` 正規化。
- `pyproject.toml` 新增 dev dependency `types-pywin32`，供 mypy 使用。
- 對應 Review Focus 第 1、2 項，測試已全數覆蓋。

## Task 6 完成內容（`src/email_downloader/client.py`）

- `OutlookClient.__init__(session_factory: ComSessionFactory | None = None)`：`protocols.py` 的 `ComSessionFactory` 從這裡開始被實際使用。
- `search()`：`resolve_folder` → `build_restrict_filter` 有值才 `Restrict()` → `Sort("[ReceivedTime]", True)` → 逐筆用 `mail_item_to_message()` 轉換、`message_matches()` 過濾、`limit` 提早中止。
- `find_latest()`：呼叫 `search(query, limit=1)` 取第一筆，沒有則回傳 `None`。

### Task 6 code review 處理結果

1. **`search()` 逐筆迴圈錯誤隔離** —— 已修正：單筆 `mail_item_to_message()` 失敗時跳過該筆、繼續處理下一筆（`except Exception:  # noqa: BLE001 -- skip unreadable item; logging deferred to a later task.`），不讓一封壞信中斷整個搜尋、遺失已收集的結果。Logging 留到 Task 9 一起補。
2. **`com_backend.py` 的 `session()` 例外遮蔽** —— 已修正：把 `yield namespace` 移到內層 try 之外，只有 `Dispatch`／`GetNamespace` 失敗才轉譯成 `OutlookConnectionError`，`with` 區塊內（`search()` 的邏輯）丟出的例外會如實傳遞。
3. **`items.Sort()` 直接呼叫在 `folder.Items`（live collection）上** —— 未修正，維持現狀：沒有 Outlook-side filter 時，`items` 還是資料夾本身，`.Sort()` 有可能改動使用者 Outlook UI 上的實際排序視圖。這是 plan 文件本身 pseudocode 就有的設計，非本次新增偏差，併入 Task 9/10 真機 smoke test 一起驗證。
4. **MailItem 的 `43` 是重複的 magic number** —— 未修正：`client.py`（`!= 43  # 43 corresponds to MailItem`）與 `com_backend.py`（`!= 43  # olMail`）各自寫一次，尚未抽成共用具名常數。低優先度小清理，待有空再處理。

## Task 7 完成內容（`src/email_downloader/attachments.py`）

- `select_attachments()`、`resolve_destination()`（`overwrite`／`skip`／`error`／`rename` 四種策略，`rename` 保留多重副檔名如 `.tar.gz`）、`safe_attachment_name()`（防路徑穿越，用 `PureWindowsPath(...).name`）。

### Task 7 意外事故：合併衝突把 `com_backend.py` 整個檔案內容複製了兩次

分支合併（commit `ee71a4f "0924 resolve conflict"`）沒有正確解決衝突，把兩份 `com_backend.py` 內容都留下，導致 `Pywin32SessionFactory`／`get_sender_email`／`_safe_getattr`／`mail_item_to_message` 全部被定義兩次。`pytest` 當時仍會過（Python 允許重複定義，後面蓋掉前面），但 `mypy` 有 4 個 `no-redef`、`ruff` 有 7 個 `F811`。**已修正**：刪掉舊的重複區塊，只留下含 Task 6 session() 修正的那份（commit `31bd13f "fix: clean up duplicate code in com_backend.py"`）。同時也順手移除了 `attachments.py` 裡沒用到的 `from itertools import count`。

**教訓**：之後解合併衝突時要逐一確認每個檔案的最終內容，不能只看有沒有衝突標記；這份 `docs/progress.md` 本身也是同一次合併的受害者（見文件開頭的提醒）。

## Task 8 完成內容（`src/email_downloader/client.py` 的 `download_attachments`／`download_latest`）

- `download_attachments()`：建立 output 目錄 → `select_attachments()` 篩選 → 沒有符合的附件丟 `AttachmentNotFoundError` → `GetItemFromID()` 重新取得信件（失敗轉 `MailAccessError`）→ 逐一取得附件、檢查檔名是否與 search 時一致（不一致轉 `MailAccessError`）→ `resolve_destination()` 決定路徑 → `SaveAsFile()`。
- `download_latest()`：`find_latest()` 找不到信丟 `MailNotFoundError`，找到就呼叫 `download_attachments()`。

### Task 8 code review 處理結果（分支 `feature/task8_attachment-download`，尚未 commit）

1. **附件被刪除時洩漏原始 pywin32 例外** —— 已修正：`item.Attachments.Item(attachment_info.index)` 包進 try/except，失敗轉成 `MailAccessError`（訊息包含 index 與信件主旨）。已補測試 `test_download_attachments_raises_when_attachment_missing`。
2. **迴圈中途某個附件存檔失敗，已成功的路徑會遺失** —— 已修正為 **best-effort**：其他附件照樣試著存，失敗的記到 `failures` 清單、`continue` 處理下一個；全部跑完後如果有失敗，才一次丟出 `AttachmentSaveError`，訊息同時列出失敗的檔名和已經成功存下來的路徑，呼叫端不會完全看不到部分成功的結果。`except Exception:` 這裡是刻意的寬例外，已加 `# noqa: BLE001` 說明原因。已補測試 `test_download_attachments_saves_remaining_when_one_fails`（兩個附件、一個成功一個失敗）。
3. **`item.Attachments` 每次迴圈都重新讀取** —— 已修正：提到迴圈外面，`attachments = item.Attachments` 只取一次。
4. **合併衝突殘留的死註解**（`# safe_name = safe_attachment_name(attachment.FileName)`）—— 已刪除。

## 待辦：`protocols.py` 已於 Task 6 開始被使用

`ComSessionFactory` Protocol 現在由 `OutlookClient.__init__` 使用，不再是死代碼。

## Task 9 完成內容（分支 `feature/task9_public-api`，commit `3450914`）

- **Public exports**：`src/email_downloader/__init__.py` 只匯出 plan 指定的穩定介面（`OutlookClient`、`MailQuery`、`MailMessage`、`AttachmentInfo`、例外 hierarchy），`Pywin32SessionFactory`、folder resolver、filter builder、COM constants 皆未匯出。`tests/units/test_public_api.py` 驗證這組 import 介面。
- **Logging**：`client.py`、`com_backend.py` 都改用 `logger = logging.getLogger(__name__)`，未呼叫 `logging.basicConfig()`，不記錄信件本文。
  - `DEBUG`：`search()` 的 query 摘要、Restrict filter 內容、命中數；`get_sender_email()` 的 Exchange SMTP fallback。
  - `INFO`：`download_attachments()` 附件成功下載、因同名檔衝突被 skip。
  - 順手修正一個小 bug：`download_attachments()` 迴圈中 `SaveAsFile()` 失敗後原本沒有 `continue`，會把存檔失敗的路徑也塞進回傳的 `paths`；補上 `continue` 後 `paths` 只會包含真正存檔成功的路徑，與 `AttachmentSaveError` 訊息裡列出的「已成功路徑」一致。
- **README**：從空檔案補齊完整內容——前置條件、三種 `uv add` 安裝方式（editable／git／registry）、`search()`／`find_latest()`／`download_attachments()`／`download_latest()` 範例、四種 conflict policy 行為表、例外處理範例（含 `MailNotFoundError` 對應「今日尚未收到資料」的業務情境）、package 與呼叫端的責任邊界、worker thread COM lifecycle 已由 package 內部處理的說明。
- **mypy strict 補洞（`uv run mypy tests`）**：先前 `[tool.mypy]` 只設定 `packages = ["email_downloader"]`（只查 `src`），這次額外把 `tests/` 也跑過 mypy strict，修正 10 個既有錯誤（非本次 diff 新增，屬於之前 commit 留下的技術債）：
  - `test_model.py`：`MailQuery(attachment_extensions=[...])` 傳 `list` 改成 `tuple`，符合欄位型別。
  - `test_com_backend.py`：動態組 fake `pythoncom`／`win32com` module 並賦值屬性的 4 處，加 `# type: ignore[attr-defined]`（mypy 官方對這種動態 module fake 手法的建議寫法）。
  - `test_client.py`：`FakeItems.__iter__`、`FakeSessionFactory.session` 補上回傳型別註解（`Iterator[Any]`／`Iterator[object]`）；3 處直接對 instance method 賦值做 mock（`client.find_latest = Mock(...)` 等）加 `# type: ignore[method-assign]`。

## 最後驗證狀態（2026-09-25，Task 9 完成並已 commit）

- `uv run pytest tests/units -v`：68 passed
- `uv run ruff check .`：All checks passed
- `uv run mypy src`：Success（9 source files）
- `uv run mypy tests`：Success（8 source files）
- Commit：`3450914 0925 feat: finish readme, logging, public api`（分支 `feature/task9_public-api`，working tree 乾淨）

## Task 12 完成內容（commit `bbb01d0`，PR #2 已合併至 main）

- `MailQuery.recursive: bool = False`：新欄位，不用額外驗證邏輯。
- `folders.iter_folder_tree(folder: Any) -> Iterator[Any]`：DFS 走訪一個資料夾自己與底下所有層級的子資料夾，不限深度。
- `client.py` 新增模組層級私有函式 `_search_folder(folder, query, *, limit)`：把原本 `search()` 裡「搜一個資料夾」的邏輯抽出來，`recursive=False`／`True` 都會用到。
- `search()` 分支邏輯：
  - `recursive=False`（預設）：只呼叫一次 `_search_folder(root_folder, query, limit=limit)`，行為與提早中斷優化跟改之前完全一樣。
  - `recursive=True`：對 `iter_folder_tree(root_folder)` 的每個資料夾呼叫 `_search_folder(folder, query, limit=None)`（不提早中斷）、`results.extend(...)`；某個子資料夾整個處理過程中丟例外就跳過（`except Exception:  # noqa: BLE001, S112 -- ...`，記一筆 `logger.warning`），不影響其他資料夾；全部資料夾跑完後才對 `results` 依 `received_time` 做一次全域排序，最後才用 `limit` 截斷。
  - 只有 `query.folder` 這個根路徑本身不存在時（`resolve_folder()` 丟出的 `FolderNotFoundError`），才維持原本行為，不會被上面的 subfolder try/except 蓋掉。
- `find_latest()`／`download_latest()` 不用改介面，`recursive` 隨著 `query` 物件自動帶過去。

### Task 12 開發中的一次 TDD 除錯紀錄

照計畫規定「先寫測試」，測試（`iter_folder_tree` 的多層走訪、`search(recursive=True)` 的合併排序／limit 延後截斷／子資料夾失敗跳過）先寫好、跑起來全部紅燈後才貼實作。第一版實作貼上後除了 3 個小 import／型別標註問題（`ruff --fix` 自動修、`noqa` 少蓋 `S112`、mypy strict 缺回傳型別），還抓到一個**所有 `search()` 測試（含完全沒用到 recursive 的舊測試）全部回傳空 list** 的真實 bug：把原本該放在 `search()` 遞迴分支結尾的「排序＋`limit` 截斷＋log」那段程式碼誤貼進 `_search_folder()` 的 for 迴圈裡面，而且漏掉了 `results.append(message)`，導致函式在處理到第一筆符合的信件時，還沒 append 就直接 `return` 一個空的 `results`。修法：把 `results.append(message)` 補回、`return results` 移到迴圈外面，`_search_folder()` 只保留單一資料夾的搜尋邏輯，排序／截斷邏輯留在 `search()` 裡。這次是靠先寫好的測試立刻抓到，沒有這批測試的話，這個 bug 會讓 `search()`／`find_latest()`／`download_latest()` 全部靜默回傳空結果，非常危險。

## 最後驗證狀態（2026-09-26，Task 12 完成並已合併至 main）

- `uv run pytest tests/ -v`：72 passed
- `uv run ruff check .`：All checks passed
- `uv run mypy src`：Success（9 source files）
- `uv run mypy tests`：Success（8 source files）
- Commit：`bbb01d0 0926 feat: support recursive subfolder search`（PR #2，`feature/task12_recursive-folder-search` → `main` `696cf9f`），working tree 乾淨

## Task 10 完成內容（commit `d4cd806`，分支 `feature/task10_smoke-test`）

- 新增 `tests/integration/test_outlook_smoke.py`：兩個 opt-in 測試，預設用 `pytest.mark.skipif` 跳過（`RUN_OUTLOOK_INTEGRATION` 沒設成 `"1"` 就不會跑），CI（`ubuntu-latest`）不會設這個環境變數，自動跳過不受影響。
  - `test_can_search_recent_inbox_mail`：搜近 24 小時內的信件，驗證 COM session、`resolve_folder`（預設 Inbox）、`Items.Restrict` 的整條路徑。
  - `test_can_find_and_optionally_download_test_message`：`find_latest()` 找一封測試信（可用 `OUTLOOK_TEST_FOLDER`／`OUTLOOK_TEST_SUBJECT_CONTAINS` 指定專用測試資料夾）、印出 subject／sender／附件檔名；再額外設 `RUN_OUTLOOK_DOWNLOAD_INTEGRATION=1` 才會呼叫 `download_attachments()` 存到 `tmp_path`，不會動到原始信件。
- `pyproject.toml` 的 `ruff.lint.ignore` 加上 `DTZ005`（`datetime.now()` 沒帶 tz），理由跟先前已忽略的 `DTZ001` 一樣：這個 package 設計上就是要用 naive 本機時間。
- **README** 補上「開發者：在真機 Outlook 上跑 opt-in smoke test」段落，並在 `search()` 說明加一句 `MailQuery.recursive` 的介紹；同時修正「v1 範圍之外」清單裡過期的「不支援遞迴搜尋所有子資料夾」（Task 12 已經支援了）。

### 真機驗證結果（2026-09-26，Windows 11 + Classic Outlook Desktop）

- `RUN_OUTLOOK_INTEGRATION=1` 下兩個測試都 PASS，包含 `RUN_OUTLOOK_DOWNLOAD_INTEGRATION=1` 之後的實際附件下載也成功（找到一封收件匣裡的真實業務信件並下載了它的附件到暫存資料夾）。
- 這代表 `filters.py` 的 `format_outlook_datetime()`（`MM/DD/YYYY hh:mm AM/PM`）在這台機器的 Outlook locale 下可以被正確解析，Task 4 code review 記錄的「地區設定風險」在這台機器上已驗證無虞。
- Step 3（用 Outlook UI 已知信件交叉比對 `received_after` 篩選語意）：待確認是否已手動執行。

### 最後驗證狀態（2026-09-26，Task 10 完成）

- `uv run pytest tests/ -v`：72 passed（純邏輯測試，不含 opt-in 真機測試）
- `RUN_OUTLOOK_INTEGRATION=1 RUN_OUTLOOK_DOWNLOAD_INTEGRATION=1 uv run pytest tests/integration/ -v -s`：2 passed（真機）
- `uv run ruff check .`：All checks passed
- `uv run mypy tests`：Success
- Commit：`d4cd806 0926 test: add opt-in outlook smoke coverage`（分支 `feature/task10_smoke-test`）
- 待 commit：`README.md`、`docs/progress.md` 的說明更新（尚未加進上面那個 commit）

## Task 11 進度（分支 `feature/task11_build-test`）

### Step 1：非實機測試 + coverage —— 完成

跑之前發現 `filters.py` 只有 84% coverage（低於 plan 要求的 90%），漏掉的都是 `message_matches()`／`attachment_matches()` 裡「exact／contains 條件不符合時回傳 `False`」的負向分支——現有測試大多只驗證「符合時回傳 `True`」，負向路徑（正確排除不符合的信）反而沒測到。這些是核心比對邏輯，補了 8 個測試涵蓋所有負向分支後 `filters.py` 到 100%，整體 coverage 97%（80 passed）。

### Step 2：靜態檢查 —— 完成

`ruff format --check .` 第一次跑列出 19 個檔案要重排版（這個專案從沒跑過 `ruff format`，純格式差異，不含邏輯變動；連 `docs/email-downloader-PLAN.md` 內嵌的 Python code block 都被一起格式化）。套用 `ruff format .` 後，`ruff format --check .`／`ruff check .`／`mypy src`／`pytest` 全部確認過一遍，沒有因為重排版而壞掉任何東西。

### Step 3：建 wheel + `twine check` —— 完成

`uv build` 產生 `dist/email_downloader-0.1.0-py3-none-any.whl` 與 `.tar.gz`，`uvx twine check dist\*` 兩個都 PASSED。

### Step 4：獨立 uv project 試裝 —— 完成

在 repo 之外的暫存目錄（`$env:TEMP\email-downloader-consumer`）建全新 uv project，`uv add` 剛剛建好的 `.whl` 檔案，`uv run python -c "from email_downloader import OutlookClient, MailQuery; print(MailQuery())"` 成功印出預設值（含 `recursive=False`），確認不需要 `--editable`、不需要動 `PYTHONPATH` 就能正常使用。

### Step 5／6：待決定

- Step 5（挑一個真實外倉／ICS 專案做最小整合，連續執行兩次驗證 conflict policy）：需要使用者自己在其他專案操作，目前尚未進行。
- Step 6（打 `v0.1.0` tag）：建議等 Step 5 實際跑過、確認沒問題後再打，避免打了 tag 之後 Step 5 又發現要改的地方。

### 最後驗證狀態（2026-09-26，Task 11 Step 1-4 完成，尚未 commit）

- `uv run pytest -m "not outlook_integration" --cov=email_downloader --cov-report=term-missing`：80 passed，整體 coverage 97%
- `uv run ruff format --check .`：22 files already formatted
- `uv run ruff check .`：All checks passed
- `uv run mypy src`：Success（9 source files）
- `uv build` + `uvx twine check dist\*`：wheel／sdist 都 PASSED
- 獨立 consumer project 試裝：import 成功
- 待 commit 檔案：`tests/units/test_filters.py`（新增 8 個測試）、以及 `ruff format .` 重排版的 19 個檔案

## Task 11 Step 5 期間發現的效能問題與修正：`build_restrict_filter()` 改用 DASL（2026-09-29）

**背景**：在 `D:\Code\test_email`（用 `uv add --editable` 裝 `email-downloader`）做 Task 11 Step 5 的真實情境測試——搜尋「憶鼎」兩個寄件者（`yidin2026@gmail.com`、`yidin2028@gmail.com`）9/1 至今、主旨含「庫存表」、附件含「庫存回饋表」、`.xlsx` 的庫存表信件，`folder="收件匣"` 加 `recursive=True`。

**發現的問題**：兩個寄件者各自 `search()` 花了 100.48s／147.21s，合計快 4 分鐘。

**除錯過程**：懷疑過是 `items.Sort()` 排序整個信件匣造成的（使用者提出的假設），寫了 `diagnose_search.py` 把 `_search_folder()` 拆成「拿 `Items`」「`Restrict()`」「`Sort()`」「逐筆轉換＋比對」四個階段分開計時，逐一資料夾量測，結果推翻了 Sort 假設：

| 資料夾 | Restrict 後筆數 | 轉換筆數 | 符合筆數 | 花費時間 |
|---|---|---|---|---|
| 收件匣（根目錄） | 1361 | 1306 | 0 | 36.13s |
| MSS System | 2501 | 2501 | 0 | 34.48s |
| MSS | 152 | 151 | 0 | 24.04s |
| 統昶 | 172 | 172 | 0 | 6.58s |
| 憶鼎（真正的資料夾） | 50 | 50 | 14 | 0.71s |

`get_items`／`Restrict`／`Sort` 三階段每個資料夾都是 0.00~0.01s，100% 的時間都花在最後「逐筆轉換＋比對」。**根本原因**：`build_restrict_filter()` 當時只把 `ReceivedTime`／`HasAttachment`／`UnRead` 推給 Outlook 端的 `Items.Restrict()`，`sender`／`subject` 完全沒有被推下去，導致每個資料夾裡「日期範圍內但跟憶鼎無關」的信，都要完整轉換成 `MailMessage`（`EntryID`／`Subject`／`SenderEmailAddress`／附件清單...每個屬性一次 COM round-trip）之後，才能在 Python 端的 `message_matches()` 被刷掉——白白轉換了近 4000 封無關信件。

**修正方案**：把 `build_restrict_filter()` 整個改寫成 DASL（`@SQL=(...)`）語法，`sender`／`sender_contains`／`subject`／`subject_contains` 也推給 Outlook 端（`urn:schemas:httpmail:fromemail`／`urn:schemas:httpmail:subject`，exact 用 `=`、contains 用 `LIKE '%...%'`）。

- Outlook 的 `Items.Restrict()` 一次呼叫裡，`@SQL=` DASL 語法跟原本 `[PropertyName] = value` 方括號語法**不能混用**，所以是整個函式改寫，不是加在旁邊。
- 新增 `_escape_dasl_literal()` 處理字串裡的單引號（`'` → `''`）。
- 附件相關條件（`attachment_name`／`attachment_name_contains`／`attachment_extensions`）**維持 Python 端過濾**：`Items.Restrict()` 只能篩 `MailItem` 本身的屬性，碰不到附件子集合，DASL 也一樣做不到。
- `message_matches()`／`attachment_matches()` **完全沒改**，繼續在 Python 端把每封通過 Outlook 端篩選的信再驗證一次，當作正確性防線——即使 DASL `LIKE` 的大小寫或萬用字元行為跟預期有落差，也不會讓錯誤的信被誤判成符合，純粹是效能優化、不改變任何行為保證。`MailQuery` 的公開 API 完全沒變。

**真機驗證結果（2026-09-29）**：改完之後在 `test_email` 重跑同一組查詢，兩個寄件者合計從 ~250 秒降到 **0.63s + 0.39s ≈ 1 秒**，結果正確——原本的 19 封都還在，多了當天（2026-09-29）新收到的 1 封（合理的新資料，不是 bug），共 20 封。

**已知殘留風險（尚未真機驗證，類比 Task 4 的 `format_outlook_datetime` locale 風險，先實作、之後真機驗證、不行再調整）**：

- `has_attachment`／`unread_only` 這兩個 DASL 布林值子句（`hasattachment = true/false`、`read = false`）這次真機測試沒有實際觸發到，只有 sender/subject/date 被驗證過。
- DASL `LIKE` 用 `%`／`_` 當萬用字元，目前 `_escape_dasl_literal()` 只跳脫單引號，沒有跳脫這兩個字元；如果 sender/subject 內容剛好包含它們，比對行為可能不如預期（機率極低的邊界情況，暫不處理）。

**TDD 執行紀錄**：`test_builds_received_and_boolean_filter` 舊斷言（方括號格式）已更新成 DASL 格式；新增 6 個測試涵蓋 exact sender／sender LIKE／exact subject／subject LIKE／單引號跳脫／`has_attachment=False`。

### 最後驗證狀態（2026-09-29，DASL 改寫完成，尚未 commit）

- `uv run pytest -m "not outlook_integration" --cov=email_downloader --cov-report=term-missing`：86 passed，`filters.py` 100% coverage，整體 coverage 97%
- `uv run ruff format --check .` / `ruff check .`：All checks passed
- `uv run mypy src` / `mypy tests`：Success（9 source files）
- 真機驗證（`D:\Code\test_email`，Windows 11 + Classic Outlook Desktop）：搜尋耗時從 ~250s 降到 ~1s，結果正確
- 待 commit：`src/email_downloader/filters.py`、`tests/units/test_filters.py`
