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

## R-IMG 圖片內嵌與轉檔器建置

來源:Gitea semi/opensrcOper #20 之審計 md2docx-run01(2026-10-29,路徑方案:w-md2html 改以本機路徑引用圖,由本套件之轉檔器讀檔內嵌;三套件 w-html2docx,w-md2html 先發,w-md2docx 最後).

| 規則 | 內容 | 站點 | 測試 |
| --- | --- | --- | --- |
| R-IMG-01 | `FormattedText`複製進模板後,對`InlineShapes`與`Shapes`(浮動圖片)之連結圖設`LinkFormat.SavePictureWithDocument=True`並`BreakLink()`.輸出 docx 之圖片關聯**無**`TargetMode="External"`,`<a:blip>`/`<asvg:svgBlip>`為`r:embed`(無`r:link`),`word/media/`有對應圖(png/jpg 位元組同來源;svg 為向量加 png 預覽,Word 會重寫 svg 內容) | `srcPython/htmlToDocx.py`之`embedLinkedPictures`(唯一調用`BreakLink`處) | `WHtml2docx.test`〈local pictures〉 |
| R-IMG-02 | 斷開後仍可讀`LinkFormat.SourceFullName`者(Word 解析不到之來源:圖檔不存在,網址取不到,正斜線絕對路徑,百分比編碼之相對路徑與 file 網址)→ 在另存前拋`LinkedPictureError`:stderr 每個來源一行`unresolved linked picture: <SourceFullName>`,stdout`error: LinkedPictureError: N linked picture(s) ...`,離開碼 1,不產生輸出檔;`WHtml2docx`沿用`code[1]:\n...`字串 reject(含清單),不改 JS 行為 | `htmlToDocx.py`之`htmlToDocxCore`、`core` | 同 |
| R-IMG-03 | 判斷[仍為連結圖]**只能**以`LinkFormat.SourceFullName`可讀為準:非連結圖取`s.LinkFormat`亦回物件不拋例外(曾因此 10 案全誤判);不以`Type`值判斷(審計實測 SVG 之連結圖 Type 為 18 而非 4,斷開後為 17) | `getLinkedPictures` | 變異 M2,M5(見報告) |
| R-IMG-04 | data URI 圖片行為不變(圖數,顯示尺寸:320×180 px PNG→240×135 pt,240×120 px JPEG→180×90 pt) | — | 同 |
| R-IMG-05 | 本機路徑 Word 可解析:反斜線絕對路徑,相對路徑原字(相對 html 所在資料夾),`file:///`原字;不可:正斜線絕對路徑,百分比編碼之相對路徑,百分比編碼之 file 網址.http(s)網址由 Word 取回內嵌.寫於 README〈圖片〉 | README | 同 |
| R-EXE-01 | 轉檔器建置輸出為`bin/htmlToDocx.exe`(`srcPython/build.bat`:`python -m PyInstaller htmlToDocx.spec --noconfirm --clean --distpath ..\bin`);`src/htmlToDocx.exe`(S2 過渡舊路徑檔)不更新.改 python 後須重建 exe,並依 R-BIN-08 兩步提交更新清單 | `srcPython/build.bat` | `unit-binManifest`(T2) |

盤點指令:

```
git grep -n "BreakLink\|SavePictureWithDocument" -- srcPython
git grep -n "LinkFormat" -- srcPython
```

判準:輸出 docx 之`word/_rels/document.xml.rels`中類型為 image 之關聯無`TargetMode="External"`;`BreakLink`只出現在`embedLinkedPictures`.

改動範圍 → 測試:改`srcPython/htmlToDocx.py`→ 重建 exe → 清單同步後跑`WHtml2docx.test`(含〈local pictures〉)與`unit-binManifest`;本套件無瀏覽器 e2e.

## R-TEST 測試

- 需實際`htmlToDocx.exe`之功能測試(`WHtml2docx.test`)以`test/tools/binEnv.mjs`之`setBinEnv(tag)`取得:將`bin/`實檔放入`test/_tmp/<tag>/mirror/`,設`W_BIN_MIRROR`,`W_BIN_CACHE`,`W_BIN_OFFLINE=1`,測完還原並刪除,**不連外網,不寫`bin-cache/`**.
- `api-ensureBins`以本機 HTTP 伺服器(port 8937)扮`raw.githubusercontent.com`(注入`fetch`改指本機,其餘網址一律拒絕),不連外網.
