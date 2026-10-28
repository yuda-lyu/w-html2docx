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
