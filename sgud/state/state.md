# 狀態

本檔記載當前已開發功能之最新狀態(只記現況,過程寫入`sgud/exp/experience.md`),與`sgud/rule/rulebook.md`(藍圖)比對即知尚未做到者.接續工作時先讀本檔.

最近更新:2026-10-29(html2docx-run01).`package.json`之`version`未動(1.0.39);`dependencies.wsemi`為`^1.9.27`.

## 圖片內嵌(規則 R-IMG-01~05,R-EXE-01)

| 項目 | 現況 | 規則 | 測試 |
| --- | --- | --- | --- |
| 本機圖與網址圖內嵌 | `srcPython/htmlToDocx.py`之`embedLinkedPictures`於`FormattedText`複製後對`InlineShapes`與`Shapes`之連結圖設`SavePictureWithDocument=True`+`BreakLink()`;輸出 docx 無 External 圖片關聯,`a:blip`為`r:embed` | R-IMG-01 | `WHtml2docx.test`〈local pictures〉 |
| 解析不到即失敗 | 斷開後仍可讀`SourceFullName`者→`LinkedPictureError`,stderr 每來源一行,stdout`error: ...`,離開碼 1,不另存;`WHtml2docx`以`code[1]:\n...`字串 reject(JS 行為未改,僅補 JSDoc) | R-IMG-02,03 | 同 |
| data URI 圖 | 行為不變(圖數,尺寸) | R-IMG-04 | 同 |
| 路徑寫法 | README〈圖片〉與〈Breaking changes〉已寫(反斜線絕對,相對原字,`file:///`原字可;正斜線絕對,百分比編碼不可) | R-IMG-05 | — |
| 轉檔器建置 | `srcPython/build.bat`輸出`..\bin\htmlToDocx.exe`;`srcPython/說明.txt`已更新;`src/htmlToDocx.exe`不更新 | R-EXE-01 | `unit-binManifest` |

## 外部執行檔(規則 R-BIN-01~08)

| 項目 | 現況 | 規則 | 測試 |
| --- | --- | --- | --- |
| 執行檔存放 | `bin/htmlToDocx.exe`已重建(含圖片內嵌;sha256`b9d244e2ef64…`,9520499 位元組;舊為`9c763ac57b88…`,9520316),待主控提交;`src/htmlToDocx.exe`保留舊版(舊版自主幹下載,不入 npm 包,不更新,sha256`9c763ac5…`) | R-BIN-05 | `unit-binManifest` |
| 清單 | `bin/manifest.json`,pinned,commit`7f73d2c236193382ee789a97fa86cd1dcf83ce1f`,1 檔(sha256`9c763ac57b88…`)——**與`bin/htmlToDocx.exe`現不一致**(兩段式第一段:執行檔先提交,清單由主控以`genBinManifest --commit <新提交>`重產) | R-BIN-02 | `unit-binManifest`(T2) |
| 取得 | `src/ensureBins.mjs`轉呼叫 wsemi`fsEnsureBins`;`WHtml2docx`轉檔前經它取路徑;`WHtml2docx.ensureBins`供部署前預先取得;`autoDownloadFiles.mjs`為薄包裝 | R-BIN-03,06 | `api-ensureBins`,`WHtml2docx.test` |
| 已移除 | `src/downloadFiles.mjs`(`fsDownloadFile`直連),`scripts/install.mjs`及`package.json`之`postinstall` | R-BIN-03 | `unit-binManifest`(無 postinstall) |
| 忽略規則 | `.npmignore`:`/bin/*`,`!/bin/manifest.json`,`/bin-cache/`,`sgud`;`.gitignore`:`/bin-cache/` | R-BIN-01 | `unit-binManifest`(T1,S1/S5) |
| 環境變數與離線 | `W_BIN_CACHE`,`W_BIN_MIRROR`,`W_BIN_OFFLINE`;README〈外部執行檔〉已寫 | R-BIN-04,07 | `api-ensureBins`(T3~T7) |

## 驗證紀錄

- 2026-10-28:`unit-binManifest`(7),`api-ensureBins`(11)新增;`WHtml2docx.test`(20,含新增 1 案,以本機鏡像供檔,實際調用 Word 轉檔)全綠.T8(主幹模式)不適用.
- 2026-10-29(html2docx-run01):`WHtml2docx.test`新增〈local pictures〉15 案(`api-ensureBins`11,`unit-binManifest`7,`WHtml2docx.test`35,合計 53).於[清單重產後]之模擬夾(新 exe+清單改為其 sha256/size)全套 53 passing;repo 內`npm test`在清單重產前為 12 passing,31 failing,31 項皆為清單與`bin/`不符所致(T2、依清單取 exe 之`WHtml2docx.test`與`api-ensureBins`各案),重產清單後應轉綠.修前(舊 exe)跑〈local pictures〉為 13 failing;變異 M1~M7 除等價之 M4(省略`SavePictureWithDocument=True`)外皆被殺.

## 已知缺口與待辦

- **清單待重產**:主控提交並推送新`bin/htmlToDocx.exe`後,執行`node toolg/genBinManifest.mjs --commit <新提交>`重產`bin/manifest.json`,再跑全套(R-BIN-08).重產前`npm test`有預期之紅(T2,及依清單取得 exe 之案).
- `dist/w-html2docx.umd.js`仍是舊建置,發布時自動重新產製.
- `docs/`(JSDoc 靜態站)未重新產製,屬發布流程.
- 下游 w-md2docx 的`src/utils.mjs`之`exeFound`診斷以 cwd 的`src/htmlToDocx.exe`查找,本改造後執行檔在`bin-cache/`,該診斷將恆為 false(僅診斷顯示,不影響轉檔);由 w-md2docx 主責處理.w-md2html(關閉 base64 時改以絕對路徑引用圖)與 w-md2docx 為路徑方案之下游,見審計 md2docx-run01.
- 範圍外未處理:html 之表格儲存格`background`屬性之圖不在`InlineShapes`/`Shapes`內,本輪未驗證其轉檔後是否內嵌;模板本身若帶有解析不到之連結圖,現會使轉檔失敗(整份文件的連結圖皆檢查).
- 發布與清單之提交由使用者/主控進行.
