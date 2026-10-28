# 經驗

執行時之累積經驗,留給後續嘗試,執行,判斷.

## 外部執行檔改造(2026-10-28,bin-html2docx)

- 照 wsemi-run04 報告 §8 範本即可:`toolg/genBinManifest.mjs`,`src/ensureBins.mjs`原檔複製,B 類單一檔不需改設定段(`archives`,`vendors`皆空).`genBinManifest`須`--commit <含 bin/ 實檔之已推送提交>`且`bin/`須與該提交相同,否則拒絕.
- 套件主檔為 default 匯出且 rollup 打 UMD:加具名匯出會讓`require('w-html2docx')`變成`{ default, ensureBins }`(破壞既有 CJS 使用者).故以靜態屬性`WHtml2docx.ensureBins = ensureBins`匯出.
- UMD 中`import.meta.url`被 rollup 轉成`pathToFileURL(__filename)`(CJS 分支);`dist/`與`src/`同為套件根下一層,`ensureBins.mjs`之`path.resolve(dirname, '..')`在兩處皆為套件根,已以打包後的 UMD 在「`pkg/dist` + `pkg/bin/manifest.json`」配置下實跑驗證.
- 原`WHtml2docx`以字串 reject,故取得執行檔失敗時轉為`[code] message`字串;直接呼叫`ensureBins()`才是帶`code`之 Error.
- 功能測試要實際執行檔又不能連網:`test/tools/binEnv.mjs`以`bin/`實檔組本機鏡像並設三個環境變數(離線),測完還原.快取在`test/_tmp/`,不污染`bin-cache/`.
- 離線缺檔測試對 wsemi 的記憶(同行程已校驗且檔案識別未變者不重雜湊):竄改快取檔後須讓 mtime 改變(`fs.utimesSync`)才能穩定觸發重驗,測試已加.
- `npm i wsemi@1.9.27`(Verdaccio)後`package-lock.json`會大幅變動(去掉`resolved`欄位,`version`同步為 package.json 的 1.0.38),屬本機 npm 設定之副作用,非本輪邏輯變更.
- 測試鎖兩槽曾被他套件(wsemi)占用,`--lock --lock-wait`等約 20 秒即取得;等待期間不要改會被排隊中指令載入之檔案.

## 本機圖片內嵌(2026-10-29,html2docx-run01,審計 md2docx-run01 之路徑方案)

- Word 匯入 html 時本機圖與 http 圖一律成連結圖(`LinkFormat.SavePictureWithDocument=False`),`SaveAs2`存為`TargetMode="External"`且`<a:blip>`只有`r:link`;`word/media/`只有 data URI 的圖.改法:`FormattedText`複製後逐一`SavePictureWithDocument=True`+`BreakLink()`.
- **判斷[仍為連結圖]不能靠`s.LinkFormat`取得與否**:非連結圖取`s.LinkFormat`亦回物件不拋例外,只有讀`SourceFullName`才拋.首版以前者判斷,10 案全誤判為未內嵌(全部失敗,訊息來源皆(unknown)).改以`SourceFullName`可讀為準.
- `Shapes`(浮動圖片)與`InlineShapes`是兩個集合.實測 Word 匯入 html 時:`<img align="left|right">`(含`hspace`,`vspace`)成`Shapes`(Type 11 連結圖,斷開後 13,存成`wp:anchor`);`style="float:left|right"`,`position:absolute`(含包在 div 內,有無 DOCTYPE 皆同)仍是`InlineShapes`(Type 4);VML(`<v:imagedata src>`)成`Shapes`但 Type 13 且已內嵌,非連結圖.表格儲存格`background`屬性之圖不在兩集合內.**要驗`Shapes`分支須用`align`寫法**:本輪首版測試用`position:absolute`,變異驗證(只處理`InlineShapes`之變異版)竟存活,才發現該寫法根本不產生`Shapes`——測試通過但沒走到該分支.
- Word 對解析不到之來源:`BreakLink`不拋例外,圖仍為 24×24 之連結圖,`SourceFullName`為 Word 改寫後之字串(正斜線絕對路徑→`C://專案/...`,百分比編碼→`%E5%9C%96 R2.png`之類,檔名片段仍可辨識).http 圖:`BreakLink`時 Word 向伺服器取圖(單次轉檔對同一網址請求約 6 次),404 者仍為連結圖.
- Word 對內嵌圖:png,jpg 位元組不變;svg 會被 Word 重寫(320 位元組 vs 來源 284),故測試不比對 svg 位元組;位元組相同之圖(如同一 JPEG 一為檔案一為 data URI)Word 只存一份.
- **同一台機器上同時兩個調用 Word 之行程會互相干擾**(`com_error: 訊息篩選器顯示應用程式正在忙碌中`,`RPC 伺服器無法使用`,`AttributeError: Open.Range`,`Property 'Word.Application.Visible' can not be set`),因 Dispatch 共用實例、一方 Quit 殺掉另一方.`quiet-run --lock`以`--cwd`區分套件:對同一套件之探測與測試務必用**同一個 cwd(repo 根)**取鎖,在 repo 根以包裝腳本(`spawnSync(..., { cwd })`)切到他處執行;不同 cwd 之兩指令會同時取得鎖而互撞(本輪即因探測與測試以不同 cwd 取鎖而撞,結果全作廢重跑).
- 新增調用 Word 之測試併入既有`WHtml2docx.test.mjs`(檔內循序),不另開檔:mocha`--parallel`下不同檔案會同時調用 Word.
- 模擬[清單重產後]之驗證:`bin/manifest.json`之 sha256/size 在主控提交前不可改,故於`tmp/<任務>/rc/`複製`src/`,`test/`,設定檔與新 exe,清單改為新 exe 之 sha256/size 後在該夾跑 mocha(`node ../../../node_modules/mocha/bin/mocha.js`,模組自上層`node_modules`解析),repo 之清單與索引不動.
- PyInstaller 建置約 10 秒(`python -m PyInstaller htmlToDocx.spec --noconfirm --clean --distpath ..\bin`);每次建置之位元組不同(sha256 隨之不同),故測試與清單不可預期固定雜湊,每次重建 exe 後須依 R-BIN-08 重產清單.
- `LinkFormat.SavePictureWithDocument = True`在已測情境(png,jpg,svg,本機路徑,http,行內與浮動)下可有可無:變異版只呼叫`BreakLink()`(M4)全部測試仍過.依審計原型與建議檔仍保留該行(Word 文件以其為先存圖資料再斷開之語意);若日後要精簡,先以更多情境(超大圖,唯讀來源)實證再刪.
- 變異驗證:變異版 py 複製到`tmp/<任務>/mut/<名>/`連同 spec 單獨建置成 exe,以 rc 夾逐一換入跑[local pictures]測試.
