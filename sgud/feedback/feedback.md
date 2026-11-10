# 回饋

任務中發現之自有套件、全域技能、全域設定之調整與抽提提案(只增不改;處置結果另記一列並寫明承接哪一列).

| # | 日期 | 類別 | 對象 | 內容 | 依據 |
| --- | --- | --- | --- | --- | --- |
| 1 | 2026-11-10 | 自有套件 | w-docx2html,w-docx-split,w-docx2pdf | 同形缺陷(以`Dispatch`連上既有 Word,轉完`Quit`)之修正須成套三件:`DispatchEx`;收尾只關自己之文件且實例內有他人文件時不 Quit(`DispatchEx`之實例仍可被他行程之`Dispatch`連上);開輸入檔唯讀(各自一個實例後,兩轉檔同開同一檔會跳出可見之「檔案使用中」對話框而停住).驗證輔助可原檔複製本套件`test/tools/wordInstance.mjs`.他套件未改前,本套件之實例仍可能被其連上並 Quit | `sgud/rule/rulebook.md`〈R-WORD〉;`sgud/exp/experience.md`〈Word 實例隔離〉;本套件 hd408 紅跑紀錄 |
| 2 | 2026-11-10 | 全域技能 | role-command-windows(或新建 Office COM 自動化技能) | 補一條:強制結束(taskkill)開著文件之 Word/Office 後,該文件被列入`HKCU\Software\Microsoft\Office\16.0\<App>\Resiliency\DisabledItems`,此後開該檔跳出可見之「上次開啟時發生嚴重錯誤」訊息框而使自動化停住;收掉懸住之 Office 前後須檢查並只清自己造成之值.另:`DisplayAlerts=False`擋不住「檔案使用中」對話框,伺服端開檔須唯讀 | 本套件 hd408:`src/tmp.docx`被列入後每次轉檔皆停住,刪值即恢復 |
| 3 | 2026-11-10 | 全域技能 | role-command-windows §1 | Windows PowerShell 5.1 以`-File`執行無 BOM 之 UTF-8 腳本時以系統字碼頁解讀,首行中文註解可吞掉換行使`param(...)`失效(參數皆為 null,無語法錯誤提示);暫存 ps1 宜只用 ASCII 或存成帶 BOM | 本套件 hd408 看門狗腳本首次啟動之錯誤輸出 |
| 4 | 2026-11-10 | 全域技能 | 新建:Office(Word)COM 自動化之規範 | 已有 5 個自有套件以 COM 驅動 Word(w-html2docx,w-docx2html,w-docx-split,w-docx2pdf,w-md2docx),本輪踩到之機轉(Dispatch 與 DispatchEx、ROT 借用、唯讀開檔、隱藏實例仍可能跳出可見對話框、DisabledItems、擁有者檔`~$`、以 http 取圖作確定時點、可見視窗看門狗)重複性高且易錯,宜抽提為一份技能供各套件共用 | 同上 |
