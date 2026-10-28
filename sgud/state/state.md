# 狀態

本檔記載當前已開發功能之最新狀態(只記現況,過程寫入`sgud/exp/experience.md`),與`sgud/rule/rulebook.md`(藍圖)比對即知尚未做到者.接續工作時先讀本檔.

最近更新:2026-10-28(bin-html2docx).`package.json`之`version`未動(1.0.38);`dependencies.wsemi`為`^1.9.27`.

## 外部執行檔(規則 R-BIN-01~08)

| 項目 | 現況 | 規則 | 測試 |
| --- | --- | --- | --- |
| 執行檔存放 | `bin/htmlToDocx.exe`(9520316 位元組)入 git;`src/htmlToDocx.exe`保留(舊版自主幹下載,不入 npm 包) | R-BIN-05 | `unit-binManifest` |
| 清單 | `bin/manifest.json`,pinned,commit`7f73d2c236193382ee789a97fa86cd1dcf83ce1f`,1 檔(sha256`9c763ac57b88…`) | R-BIN-02 | `unit-binManifest` |
| 取得 | `src/ensureBins.mjs`轉呼叫 wsemi`fsEnsureBins`;`WHtml2docx`轉檔前經它取路徑;`WHtml2docx.ensureBins`供部署前預先取得;`autoDownloadFiles.mjs`為薄包裝 | R-BIN-03,06 | `api-ensureBins`,`WHtml2docx.test` |
| 已移除 | `src/downloadFiles.mjs`(`fsDownloadFile`直連),`scripts/install.mjs`及`package.json`之`postinstall` | R-BIN-03 | `unit-binManifest`(無 postinstall) |
| 忽略規則 | `.npmignore`:`/bin/*`,`!/bin/manifest.json`,`/bin-cache/`,`sgud`;`.gitignore`:`/bin-cache/` | R-BIN-01 | `unit-binManifest`(T1,S1/S5) |
| 環境變數與離線 | `W_BIN_CACHE`,`W_BIN_MIRROR`,`W_BIN_OFFLINE`;README〈外部執行檔〉已寫 | R-BIN-04,07 | `api-ensureBins`(T3~T7) |

## 驗證紀錄

- 2026-10-28:`unit-binManifest`(7),`api-ensureBins`(11)新增;`WHtml2docx.test`(20,含新增 1 案,以本機鏡像供檔,實際調用 Word 轉檔)全綠.T8(主幹模式)不適用.

## 已知缺口與待辦

- `dist/w-html2docx.umd.js`仍是舊建置,發布時自動重新產製;已以相同設定建置到暫存夾驗證加入`ensureBins`後可打包,UMD 載入後`ensureBins`可用且 default 仍為函數.
- `docs/`(JSDoc 靜態站)未重新產製,屬發布流程.
- 下游 w-md2docx 的`src/utils.mjs`之`exeFound`診斷以 cwd 的`src/htmlToDocx.exe`查找,本改造後執行檔在`bin-cache/`,該診斷將恆為 false(僅診斷顯示,不影響轉檔);由 w-md2docx 主責處理.
- 發布與清單之提交由使用者/主控進行;`bin/manifest.json`之`commit`引用已推送之`7f73d2c`.
