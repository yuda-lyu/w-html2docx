# 狀態

本檔記載當前已開發功能之最新狀態(只記現況,過程寫入`sgud/exp/experience.md`),與`sgud/rule/rulebook.md`(藍圖)比對即知尚未做到者.接續工作時先讀本檔.

最近更新:2026-11-10(hd408 第一段).`package.json`之`version`未動(1.0.51);`dependencies.wsemi`為`^1.9.55`.

## Word 實例(規則 R-WORD-01~05)

| 項目 | 現況 | 規則 | 測試 |
| --- | --- | --- | --- |
| 取得 Word | `srcPython/htmlToDocx.py`之`htmlToDocx`以`wc.DispatchEx('Word.Application')`另起實例(原`Dispatch`會連上已在執行之 Word) | R-WORD-01 | 〈Word instance isolation〉① |
| 收尾 | `closeOwnWord`:只關本行程之文件;`countOtherDocs`以物件身分盤點他人文件(含受保護檢視);有他人文件或無法判定時不 Quit;`Visible`/`DisplayAlerts`已移入`try` | R-WORD-02,03 | ③④ |
| 開檔 | 模板與來源唯讀開啟;`fpOut`即該輸入檔時可寫開啟(`isSamePath`) | R-WORD-04 | ②(同時,錯開 1.5 秒,交錯),⑤ |
| 執行檔 | `bin/htmlToDocx.exe`已以`srcPython/build.bat`重建(sha256`8bc802a9fb57f9cb2baed62b724a1e92a0d6949c6c3d1fbabf7ffc94e42c3917`,9521857 位元組),**待主控提交** | R-EXE-01 | 模擬夾全套 |
| 驗證法 | `test/tools/wordInstance.mjs`(可移植輔助,含可見視窗監看)+`WHtml2docx.test`〈Word instance isolation〉6 案 | — | 修前 exe 紅,新 exe 綠 |

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
| 執行檔存放 | `bin/htmlToDocx.exe`為 hd408 重建者(Word 實例隔離;sha256`8bc802a9fb57…`,9521857 位元組),待主控提交;前版`b9d244e2ef64…`(9520499,圖片內嵌,提交`c9e7ea3`);`src/htmlToDocx.exe`保留舊版(舊版自主幹下載,不入 npm 包,不更新,sha256`9c763ac5…`) | R-BIN-05 | `unit-binManifest` |
| 清單 | `bin/manifest.json`,pinned,commit`c9e7ea3131c4cdb400da36fb4dc4b7fa5f5fa864`,1 檔(sha256`b9d244e2ef64…`,9520499)——**與`bin/htmlToDocx.exe`現不一致**(兩段式第一段:執行檔先提交,清單由主控以`genBinManifest --commit <新提交>`重產) | R-BIN-02 | `unit-binManifest`(T2) |
| 取得 | `src/ensureBins.mjs`轉呼叫 wsemi`fsEnsureBins`;`WHtml2docx`轉檔前經它取路徑;`WHtml2docx.ensureBins`供部署前預先取得;`autoDownloadFiles.mjs`為薄包裝 | R-BIN-03,06 | `api-ensureBins`,`WHtml2docx.test` |
| 已移除 | `src/downloadFiles.mjs`(`fsDownloadFile`直連),`scripts/install.mjs`及`package.json`之`postinstall` | R-BIN-03 | `unit-binManifest`(無 postinstall) |
| 忽略規則 | `.npmignore`:`/bin/*`,`!/bin/manifest.json`,`/bin-cache/`,`sgud`;`.gitignore`:`/bin-cache/` | R-BIN-01 | `unit-binManifest`(T1,S1/S5) |
| 環境變數與離線 | `W_BIN_CACHE`,`W_BIN_MIRROR`,`W_BIN_OFFLINE`;README〈外部執行檔〉已寫 | R-BIN-04,07 | `api-ensureBins`(T3~T7) |

## 驗證紀錄

- 2026-10-28:`unit-binManifest`(7),`api-ensureBins`(11)新增;`WHtml2docx.test`(20,含新增 1 案,以本機鏡像供檔,實際調用 Word 轉檔)全綠.T8(主幹模式)不適用.
- 2026-10-29(html2docx-run01):`WHtml2docx.test`新增〈local pictures〉15 案(`api-ensureBins`11,`unit-binManifest`7,`WHtml2docx.test`35,合計 53).於[清單重產後]之模擬夾(新 exe+清單改為其 sha256/size)全套 53 passing;repo 內`npm test`在清單重產前為 12 passing,31 failing,31 項皆為清單與`bin/`不符所致(T2、依清單取 exe 之`WHtml2docx.test`與`api-ensureBins`各案),重產清單後應轉綠.修前(舊 exe)跑〈local pictures〉為 13 failing;變異 M1~M7 除等價之 M4(省略`SavePictureWithDocument=True`)外皆被殺.

- 2026-11-10(hd408 第一段):`WHtml2docx.test`新增〈Word instance isolation〉6 案(合計 59).修前 exe:①③②a④ 4 failing,②b 懸住並跳出可見「檔案使用中」;新 exe(`8bc802a9…`)於[清單重產後]之模擬夾全套 59 passing;repo 內`npm test`在清單重產前為 12 passing,32 failing(皆清單與`bin/`不符).新舊輸出 docx 比對差異在 Word 逐次雜訊內.外審 terra 22 條(屬實 16,部分屬實 6).

## 已知缺口與待辦

- **清單待重產(hd408 第二段)**:主控提交並推送新`bin/htmlToDocx.exe`(與`srcPython/htmlToDocx.py`,測試,sgud)後,執行`node toolg/genBinManifest.mjs --commit <新提交>`重產`bin/manifest.json`,再跑全套(R-BIN-08).重產前`npm test`有預期之紅(T2,及依清單取得 exe 之案).
- Word 實例隔離之未實測格:使用者開著可見之互動 Word 時(不得開可見視窗,依 R-WORD-01 之機轉推論不受影響);轉檔中使用者自檔案總管開 docx 是否經 DDE 進入本行程之隱藏實例(若進入,R-WORD-02 見外來文件即不 Quit);`countOtherDocs`之「無法判定→保留實例」分支無法注入,只以讀碼驗證.
- Word 實例隔離之已知限制(外審 hd408):他行程(舊寫法)先 Quit 本行程之實例,或取用本行程已開之同一文件物件時,本行程仍會失敗或關掉該文件,須對方改用`DispatchEx`(w-docx2html,w-docx-split,w-docx2pdf 由主控另派);盤點與 Quit 之間他人仍可連入;保留實例時之 stderr 提示 JS 端不轉出(`WHtml2docx`成功只回`ok`);`DispatchEx`本身於取得控制代碼前失敗時無從收尾;轉檔器被`opt.timeout`強制結束時 WINWORD 不在其行程樹(JSDoc 已載);指定模板之巨集安全性(`AutomationSecurity`未設)屬改前既有設計,未處理.
- `dist/w-html2docx.umd.js`仍是舊建置,發布時自動重新產製.
- `docs/`(JSDoc 靜態站)未重新產製,屬發布流程.
- 下游 w-md2docx 的`src/utils.mjs`之`exeFound`診斷以 cwd 的`src/htmlToDocx.exe`查找,本改造後執行檔在`bin-cache/`,該診斷將恆為 false(僅診斷顯示,不影響轉檔);由 w-md2docx 主責處理.w-md2html(關閉 base64 時改以絕對路徑引用圖)與 w-md2docx 為路徑方案之下游,見審計 md2docx-run01.
- 範圍外未處理:html 之表格儲存格`background`屬性之圖不在`InlineShapes`/`Shapes`內,本輪未驗證其轉檔後是否內嵌;模板本身若帶有解析不到之連結圖,現會使轉檔失敗(整份文件的連結圖皆檢查).
- 發布與清單之提交由使用者/主控進行.
