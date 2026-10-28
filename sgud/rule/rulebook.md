# 規則帳本(藍圖)

本檔以規則為單位記程式碼須長成什麼樣,附盤點指令與判準.與`sgud/state/state.md`(現況)比對即知尚未做到者.

## R-BIN 外部執行檔之存放、下載、校驗與離線

制式規格:`C:\opensrcOper\dspec\外部執行檔之存放下載校驗與離線.md`(本套件為 B 類改造,釘選模式,2026-10-28 bin-html2docx).**本節只摘要,細節以規格為準,不重述**;主控帳本 G10 登記站點全集.

| 規則 | 內容 | 規格條 | 測試 |
| --- | --- | --- | --- |
| R-BIN-01 | npm 包只含`bin/manifest.json`:`.npmignore`列`/bin/*`,`!/bin/manifest.json`,`/bin-cache/`,`src/htmlToDocx.exe`;`.gitignore`列`/bin-cache/`;`.gitattributes`設`bin/** binary` | S1,S5,F12 | `unit-binManifest`(T1) |
| R-BIN-02 | `bin/manifest.json`(`mode: pinned`,`commit`為含`bin/`實檔之 40 碼提交)由`toolg/genBinManifest.mjs --commit <提交>`依實檔產生,不手寫;清單與`bin/`實檔之 sha256/大小一致 | M1,M2,M4,U1 | `unit-binManifest`(T2) |
| R-BIN-03 | 取得`htmlToDocx.exe`之路徑只經`src/ensureBins.mjs`(轉呼叫 wsemi`fsEnsureBins`);**不得**自寫下載,快取,鏡像,校驗;不得在`postinstall`下載;`WHtml2docx`於轉檔前取路徑,取不到以字串 reject(含`[code]`) | D1,D2,E2,E3 | `api-ensureBins`,`WHtml2docx.test`(input checks) |
| R-BIN-04 | 環境變數`W_BIN_CACHE`,`W_BIN_MIRROR`,`W_BIN_OFFLINE`由 wsemi 統一讀取,套件端不另解析;快取預設`<套件根>/bin-cache/<sha256>/<name>`,命中亦校驗,竄改拒用 | D3,D4,O1~O3 | `api-ensureBins`(T3~T7) |
| R-BIN-05 | 舊版相容:`src/htmlToDocx.exe`保留於 repo(舊版自主幹該路徑下載;不入 npm 包),至舊版淘汰後另案移除;`src/autoDownloadFiles.mjs`保留為薄包裝(回傳`{ fpExe }`,經`ensureBins`),供舊深層匯入 | S2 | `unit-binManifest`,`api-ensureBins` |
| R-BIN-06 | 匯出`ensureBins`以靜態屬性`WHtml2docx.ensureBins`提供,**不加具名匯出**(UMD 打包後 default 形態須不變) | E2 | `api-ensureBins` |
| R-BIN-07 | README 須有〈外部執行檔〉一節(檔,首次使用下載,三個環境變數,離線部署步驟) | O5 | — |
| R-BIN-08 | 更新執行檔:替換`bin/`內檔 → 提交並推 GitHub(得提交雜湊)→`node toolg/genBinManifest.mjs --commit <該提交>` → 測試 → 發布新版;兩步提交,執行檔之提交在前 | U1 | `unit-binManifest` |

盤點指令:

```
node C:/opensrcOper/tools/scan-bins.mjs --only w-html2docx
git grep -n "fsDownloadFile\b\|fsDownloadFileVerify\|postinstall" -- src package.json
```

判準:`src/`內除`ensureBins.mjs`外不得出現下載/校驗/快取邏輯;`package.json`之`scripts`不得有`postinstall`;`npm pack --dry-run`之`bin/`下只有`manifest.json`.

## R-TEST 測試

- 需實際`htmlToDocx.exe`之功能測試(`WHtml2docx.test`)以`test/tools/binEnv.mjs`之`setBinEnv(tag)`取得:將`bin/`實檔放入`test/_tmp/<tag>/mirror/`,設`W_BIN_MIRROR`,`W_BIN_CACHE`,`W_BIN_OFFLINE=1`,測完還原並刪除,**不連外網,不寫`bin-cache/`**.
- `api-ensureBins`以本機 HTTP 伺服器(port 8937)扮`raw.githubusercontent.com`(注入`fetch`改指本機,其餘網址一律拒絕),不連外網.
