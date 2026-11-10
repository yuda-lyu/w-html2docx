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

改動範圍 → 測試:改`srcPython/htmlToDocx.py`→ 重建 exe → 清單同步後跑`WHtml2docx.test`(含〈local pictures〉與〈Word instance isolation〉)與`unit-binManifest`;本套件無瀏覽器 e2e.

## R-WORD Word 實例須自行建立、只結束自己者

來源:主控排程第 408 項(hd408,2026-11-10;源自下游 w-md2docx 之調查:兩行程同時轉檔互相關掉對方文件).同形站點另有 w-docx2html、w-docx-split、w-docx2pdf,由主控另派,本節之規則與驗證法為其範本.

| 規則 | 內容 | 站點 | 測試 |
| --- | --- | --- | --- |
| R-WORD-01 | 取得 Word 一律`win32com.client.DispatchEx('Word.Application')`(CoCreateInstanceEx,每次另起 WINWORD;實測已有其他 Word 時亦另起).**不得**用`Dispatch`:pywin32 之`Dispatch`內部先`pythoncom.connect`(GetActiveObject)連上已在執行之 Word(他行程之轉檔或使用者開啟者),轉完`Quit(0)`即以不儲存結束對方 | `htmlToDocx.py`之`htmlToDocx` | 〈Word instance isolation〉①②④ |
| R-WORD-02 | 只結束自己之實例:收尾一律經`closeOwnWord`——關閉本行程之文件(關閉失敗者仍屬本行程,以 COM 物件身分比對,不以路徑:他人可能重開同一路徑),再以`countOtherDocs`盤點實例內非本行程之文件(含不在`Documents`內之受保護檢視視窗`ProtectedViewWindows`);有他人文件**或無法判定**(讀取失敗,有界重試 10×0.5 秒後)時**不 Quit**(stderr 印`Word instance is kept since …`),實例已不存在(RPC 斷線類錯誤)則不再動作,否則`Quit(0)`.why:`DispatchEx`所建之實例仍登錄於 ROT,他行程之`Dispatch`可連上並於其中開文件(實測);寧可留下實例,不關他人之文件(原則 3).限制:盤點與 Quit 之間他人仍可連入;他人取用本行程已開之同一文件物件時,本行程之 Close 仍會關掉它(須對方改用`DispatchEx`) | `closeOwnWord`,`countOtherDocs` | ④ |
| R-WORD-03 | 例外路徑不殘留:`Visible`/`DisplayAlerts`之設定與轉檔同在`try`內,任何例外(含開檔失敗、`LinkedPictureError`)皆經`finally`之`closeOwnWord` | `htmlToDocx` | ③(a)(b)(c) |
| R-WORD-04 | 模板與來源以唯讀開啟`Documents.Open(fp, False, not isSamePath(fp, fpOut))`(位置參數依序為 FileName,ConfirmConversions,ReadOnly);轉檔器只`SaveAs2`至`fpOut`.why:各轉檔各一個實例,兩轉檔以可寫方式同開同一檔(如內建模板)時,後開者跳出「檔案使用中」對話框(可見)而停住(實測).**例外**:`fpOut`即該輸入檔(如以輸出覆寫模板)時可寫開啟——唯讀開啟則 Word 拒絕另存回原路徑(「這是唯讀檔案,無法儲存」,實測),改前可覆寫,不得改變.限制:唯讀只解此鎖檔情境,編碼、密碼、修復、安全警告等互動仍可能出現(改前即有) | `htmlToDocxCore`,`isSamePath` | ②(同時,錯開 1.5 秒,交錯,同模板同 html);⑤ |
| R-WORD-05 | 不開可見視窗:`app.Visible = False`,`app.DisplayAlerts = False`;測試中扮演他行程之 Word 亦一律隱藏 | `htmlToDocx`,`test/tools/wordInstance.mjs` | 全組 |

盤點指令:

```
git grep -n "Dispatch(\|DispatchEx(\|GetActiveObject\|\.Quit(" -- srcPython
git grep -n "Documents.Open(" -- srcPython
```

判準:`srcPython/`取得 Word 只有`wc.DispatchEx('Word.Application')`一處,無`Dispatch(`;`.Quit(`只在`closeOwnWord`;`Documents.Open(`皆以第三個位置參數`not isSamePath(fp, fpOut)`開啟.

驗證法(可移植):`test/tools/wordInstance.mjs`不含本套件之轉檔呼叫,他套件可原檔複製(前提寫於檔頭:Windows PowerShell 5.1、Word 2016 以後、Node 18.2 以後、於 Word 測試閘內執行)——`wordPids`,`isAutomationWord`,`waitWordGone`,`killNewAutomationWords`(只結束本組開始後新出現且命令列含`/Automation`者;無法區分他人於期間另起者,故須在測試閘內),`listDisabledItems`/`clearDisabledItems`(回傳實際刪除與刪除失敗者),`startHeldWord`(以 PowerShell 經 COM 另起隱藏 Word 並持有一份含未存文字之文件,扮演他行程或使用者之 Word),`startBorrower`(以 GetActiveObject 連上隱藏之 Word、新增文件並寫入唯一未存文字,扮演舊寫法之他行程;可見之 Word 不碰;Word 忙碌時以`InvokeMember`讀寫並重試),`startVisibleWordWatch`(每 200 ms 列 WINWORD 之可見頂層視窗,自動化實例出現可見視窗即結束之;各案以`visibleSince`斷言 0 筆;短於一輪之閃現可能漏看),`startHoldServer`(扣住 http 取圖之回應,作「轉檔進行中」之確定時點).各套件仿〈Word instance isolation〉寫其轉檔呼叫;輸入無法引用 http 資源者,改以轉檔期間取樣 WINWORD 判定「另起實例」,其餘斷言(他人 Word 存活且文件不變、無殘留、無可見視窗)不依賴時點.

## R-TEST 測試

- 需實際`htmlToDocx.exe`之功能測試(`WHtml2docx.test`)以`test/tools/binEnv.mjs`之`setBinEnv(tag)`取得:將`bin/`實檔放入`test/_tmp/<tag>/mirror/`,設`W_BIN_MIRROR`,`W_BIN_CACHE`,`W_BIN_OFFLINE=1`,測完還原並刪除,**不連外網,不寫`bin-cache/`**.
- `api-ensureBins`以本機 HTTP 伺服器(port 8937)扮`raw.githubusercontent.com`(注入`fetch`改指本機,其餘網址一律拒絕),不連外網.
- 〈Word instance isolation〉(6 案:①③②a②b④⑤)須在**無其他 WINWORD**時執行(有則整組略過並印出原因:修前之轉檔器會借用並結束之,且④無法確定連上本行程之實例);略過或初始化前失敗時`after`不結束任何 WINWORD.扣住點伺服器以作業系統配發埠(`listen(0)`).收尾以`killNewAutomationWords`只結束本組新出現之自動化 Word;以`clearDisabledItems`清除指向 repo 內檔案之 Word 停用項目(before 於略過判定之後一次;after 只在有強制結束時)——強制結束開著文件之 Word 後,Word 將該文件列入`HKCU\Software\Microsoft\Office\16.0\Word\Resiliency\DisabledItems`,之後開該檔即跳出可見訊息框而停住(實測).測試是否可自動刪除使用者登錄中指向 repo 之停用項目,待主控裁示(hd408 報告〈需裁示〉).
