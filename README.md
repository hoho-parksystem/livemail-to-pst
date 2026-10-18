# LiveMail EML to Outlook PST

將 Windows Live Mail 匯出的 `.eml` 資料夾樹複製到新的 Unicode PST，讓傳統版 Outlook 開啟。來源 EML 不會被修改或刪除。PowerShell 原始碼採 MIT 授權，歡迎透過 Issue、Pull Request 一起維護。

**目前狀態：實驗性版本。已做離線驗證，尚未完成真實 Outlook + Redemption 端到端驗證。請先用郵件備份的小樣本試跑，核對後再遷移全部資料。不能保證所有 EML 都能無損轉換。**

## 必要條件

- Windows 與已安裝的 **傳統版 Outlook 桌面程式**（提供 Extended MAPI）；只有新版 Outlook、Outlook 網頁版或 Mac 不適用此腳本。
- Windows PowerShell 5.1，STA 模式。PowerShell、Outlook、Redemption 必須同為 32 位元或同為 64 位元。請在 Outlook「檔案 → Office 帳戶 → 關於 Outlook」查看位元數。
- 從 [Redemption 官方網站](https://www.dimastr.com/redemption/)取得並依官方文件註冊 COM DLL。請使用近期版本，閱讀 [修正紀錄](https://www.dimastr.com/redemption/history.htm)與授權條款。它是獨立的第三方元件，**本專案的 MIT 授權不涵蓋 Redemption 或 Outlook**；商業使用／散布請向廠商確認授權。本專案不附帶 DLL、不會自動安裝、購買或接受其條款。
- 本機可寫入的輸出目錄與足夠磁碟空間。避免網路磁碟、同步資料夾；大量郵件分批輸出不同 PST，留意 Outlook PST 大小限制。

如需手動註冊 DLL，請依廠商指示使用相符位元數的 `regsvr32`；勿從不明來源下載 DLL。通常只有元件註冊需要系統管理員權限，匯入時以一般使用者執行。

## 使用方式

1. 備份完整 Live Mail 郵件資料。使用 Live Mail 的郵件匯出功能取得 EML；此工具不讀取 Live Mail 資料庫，也不搬移帳號、密碼、聯絡人、行事曆或郵件規則。
2. 下載此儲存庫（Code → Download ZIP）並解壓縮。關閉 Outlook，遷移期間不要修改來源 EML。
3. 開啟與 Outlook 相符位元數的 Windows PowerShell，切換至腳本所在目錄。先掃描：

```powershell
.\eml_to_outlook.ps1 -EmlFolderPath 'D:\LiveMail Export' -PstOutputPath 'D:\Mail Archive\LiveMail.pst' -ScanOnly
```

掃描不建立 PST／記錄，也不測試 COM 安裝。輸出目錄必須事先存在；來源不可包含 junction／符號連結；無法讀取的資料夾會中止，不會悄悄略過。

4. 正式轉換（每次使用全新的 PST 檔名）：

```powershell
powershell.exe -NoProfile -STA -File .\eml_to_outlook.ps1 `
  -EmlFolderPath 'D:\LiveMail Export' `
  -PstOutputPath 'D:\Mail Archive\LiveMail.pst'
$LASTEXITCODE
```

64 位元 Windows 上，64 位元 Windows PowerShell 位於 `%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe`，32 位元位於 `%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe`。請由相符位元數的終端機啟動；系統的檔案重新導向可能影響從 32 位元程式呼叫 System32。

若 Windows 封鎖下載的腳本，先檢查內容，再依組織政策解除封鎖；不要關閉全機安全政策。

若是個人電腦的執行政策阻擋、且你已閱讀並信任腳本，可在上述單次 `powershell.exe` 命令加入 `-ExecutionPolicy Bypass`；這只作用於該程序，不修改全機政策。組織管控的電腦請遵循管理政策。

5. 完成後，在傳統版 Outlook 選擇「檔案 → 開啟和匯出 → 開啟 Outlook 資料檔」，選擇產生的 PST。程式使用獨立暫時 MAPI profile，不會將 PST 永久掛入你的既有 Outlook profile。

## 參數與結果

| 參數 | 說明 |
| --- | --- |
| `EmlFolderPath` | EML 根目錄，包含所有子資料夾、隱藏 EML；路徑中的 `[]` 以文字處理 |
| `PstOutputPath` | 全新 `.pst` 路徑；拒絕既有檔案，禁止放在來源內 |
| `TempFolderName` | PST 中的匯入根資料夾，預設 `LiveMail Import`；保留其下來源階層 |
| `LogPath` | 預設 `<PST 路徑>.import.jsonl`；必須全新且在來源以外 |
| `ScanOnly` | 僅列舉與檢查路徑，不匯入 |

來源根目錄的 EML 直接放在匯入根資料夾；只建立含郵件的資料夾階層，不搬移空資料夾。同一內容若出現在多個檔案，仍逐檔匯入，不去重。檔案順序不代表郵件時間順序。

退出碼：`0` = 全部儲存並可重新開啟（或掃描完成），`1` = 至少一封匯入失敗，`2` = 前置檢查／系統／記錄／關閉 MAPI 失敗。**儲存並可重新開啟不代表內容已通過無損驗證。**

JSON Lines 記錄包含開始資料、每封來源相對路徑／SHA-256／狀態／EntryID／錯誤，以及正常結束時的統計。每封完成後立即 flush。沒有 summary 表示未正常完成，應檢查部分 PST。記錄包含路徑與錯誤資訊，可能涉及隱私，請勿公開上傳。

```powershell
$rows = Get-Content -LiteralPath 'D:\Mail Archive\LiveMail.pst.import.jsonl' | ConvertFrom-Json
$rows | Where-Object { $_.type -eq 'message' -and $_.status -eq 'failed' }
```

同一輸出旁會留下空的 `.import.lock` 檔，供程式排他鎖定；它不包含郵件。中斷後不會自動刪除部分 PST，也**不支援續傳**。保留舊結果供核對，以新檔名重跑；請勿合併重跑結果而造成重複郵件。

## 核對與限制

- 比對來源封數、記錄的成功／失敗數，以及 PST 各資料夾的郵件數。失敗可能留下部分項目，請個別檢查。
- 抽查中文／Big5／UTF-8 主旨與正文、HTML／純文字、寄件者及收件者、寄送／接收日期與時區、一般附件及內嵌圖片。附件另存後比對 SHA-256；重要資料應逐封核對。
- MIME 轉換由 Redemption 負責；損壞或非標準 MIME 可能失敗或轉換不完整。程式不自行重建正文，不從檔案修改時間猜測郵件日期。
- 已讀狀態、草稿／寄件備份狀態、旗標、分類等 Live Mail 資料庫資訊未必存在 EML 中，不能保證保留。請特別檢查草稿與寄件備份。
- 加密信需要原本憑證／私鑰；簽章、特殊郵件類別與大型附件須另外驗證。不下載遠端圖片；工具沒有呼叫寄信功能。
- 來源可能含尚未同步的信件，請先確認 Live Mail 匯出完整。驗收完成並另存備份前，勿移除舊郵件或舊電腦資料。

## 為什麼重寫原本流程？

Microsoft 的 [OpenSharedItem 文件](https://learn.microsoft.com/en-us/office/vba/outlook/how-to/sharing/import-saved-items-using-openshareditem)只列出 ICS、VCF、MSG；直接傳入 EML 或改成 file URL 不是可靠的 EML 匯入方式。透過檔案關聯開啟 EML，再讀取 ActiveInspector，無法確認抓到的是否為該封信，且可能啟動其他郵件程式。

新版改用 [RDOMail.Import](https://www.dimastr.com/redemption/rdomail.htm) 與 [olRFC822_Redemption (1031)](https://www.dimastr.com/redemption/enum_rdoSaveAsType.htm)，並以 [LogonPstStore](https://www.dimastr.com/redemption/rdo/rdosession.htm) 建立獨立的 Unicode PST 工作階段。沒有以視窗焦點為依據的備援路徑。成功數只在儲存後重新開啟成功才增加，失敗獨立計算。

## 開發與共同維護

請閱讀 [CONTRIBUTING.md](CONTRIBUTING.md)；離線測試：

```powershell
powershell.exe -NoProfile -File .\tests\Test-Offline.ps1
```

測試不會匯入私人郵件。實機驗證步驟見 [測試清單](docs/VALIDATION.md)。GitHub Actions 只執行離線測試，不代表已通過 Outlook 實機驗證。

English: Experimental Windows Live Mail EML-tree to Unicode PST importer. Requires classic Outlook, Windows PowerShell 5.1 and separately licensed/registered Redemption, all with matching bitness. Creates a fresh PST, preserves relative folders, never changes source EML, and writes per-file JSONL results. No resume, no deduplication, no lossless guarantee. Validate a backup sample before migration. Contributions and sanitized bug reports welcome; do not upload private email.
