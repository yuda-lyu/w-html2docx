from win32com import client as wc
import sys

#使用win32com
#win32com因更新問題會出現 ImportError: DLL load failed while importing win32api: 找不到指定的模組。
#1. 安裝 pip install pywin32
#2. 使用系統管理員權限開啟cmd
#3. cd至安裝目錄 C:\ProgramData\Anaconda3\Scripts
#4. 用python安裝腳本 python pywin32_postinstall.py -install

#編譯
#1. 安裝編譯套件 pip install pyinstaller
#2. 編譯 pyinstaller -F htmlToDocx.py

#win32com教學
#https://zhuanlan.zhihu.com/p/67543981


def getError():
    import sys

    #exc_info
    type, message, traceback = sys.exc_info()

    #es
    es=[]
    while traceback:
        e={
            'name':traceback.tb_frame.f_code.co_name,
            'filename':traceback.tb_frame.f_code.co_filename,
        }
        es.append(e)
        traceback = traceback.tb_next

    #err
    err={
        'type':type,
        'message':message,
        'traceback':es,
    }

    return err


def j2o(v):
    #json轉物件
    import json
    return json.loads(v)


def o2j(v):
    #物件轉json
    import json
    return json.dumps(v, ensure_ascii=False)


def str2b64(v):
    #字串轉base64字串
    import base64
    v=base64.b64encode(v.encode('utf-8'))
    return str(v,'utf-8')


def b642str(v):
    #base64字串轉字串
    import base64
    return base64.b64decode(v).decode('utf-8')


def readText(fn):
    #讀取檔案fn內文字
    import codecs
    with codecs.open(fn,'r',encoding='utf8') as f:
        return f.read()


def writeText(fn,str):
    #寫出文字str至檔案fn
    import codecs
    with codecs.open(fn,'w',encoding='utf8') as f:
        f.write(str)


def getSectGrids(fpTemp):
    #讀模板各節之文件格線, 依文件順序回傳[{type, pitch}], 與Word之Sections索引同序; 讀取失敗回傳[]
    #why: Word COM之PageSetup無行距(LinePitch)屬性, 而版心高÷LinesPage並非實際行距(實測行距18點之模板為18.37), 須自模板之docGrid取得
    import re
    import zipfile

    try:
        with zipfile.ZipFile(fpTemp) as z:
            xml = z.read('word/document.xml').decode('utf-8')
    except:
        return []

    #修訂紀錄內之舊設定不計
    xml = re.sub(r'<w:(sectPrChange|pPrChange)\b[\s\S]*?</w:\1>', '', xml)

    rs = []
    for m in re.finditer(r'<w:sectPr\b[^>]*?/>|<w:sectPr\b[^>]*>[\s\S]*?</w:sectPr>', xml):
        t = 'default' #未給type即無格線
        pitch = 0
        g = re.search(r'<w:docGrid\b([^>]*)>', m.group(0))
        if g:
            mt = re.search(r'w:type="([^"]+)"', g.group(1))
            mp = re.search(r'w:linePitch="(\d+)"', g.group(1))
            if mt:
                t = mt.group(1)
            if mp:
                pitch = int(mp.group(1)) / 20 #twip轉點
        rs.append({'type': t, 'pitch': pitch})

    return rs


def getImgHeightMaxByGrid(ps, grid, contentH, reserveLines):
    #行格線模板之圖片最大高度: (每頁格數-預留格數)×行距-0.5點, 不適用時回傳None
    #實測(Word 16.0.14334): 行距即docGrid之linePitch; 每頁格數為floor(版心高÷行距); 行內圖片段佔ceil(圖高÷行距)格, 不另加下沿;
    #  由k格進位為k+1格之門檻多恰為k×行距, 部分行距(如20點)下提前至多0.25點, 故減0.5點
    import math

    if grid is None or grid['type'] not in ('lines', 'linesAndChars', 'snapToChars'):
        return None
    if ps.LayoutMode not in (1, 2, 3): #wdLayoutModeGrid, wdLayoutModeLineGrid, wdLayoutModeGenko
        return None
    pitch = grid['pitch']
    if pitch <= 0:
        return None
    nLines = int(math.floor(contentH / pitch + 1e-6))
    if reserveLines >= nLines:
        print(f'imgHeightReserveLines[{reserveLines}] is not less than the lines per page[{nLines}], use the default max height', file=sys.stderr)
        return None
    return (nLines - reserveLines) * pitch - 0.5


def getImgSizeMax(ps, grid, opt):
    #圖片最大寬高, 依所在節之版心計算

    #版心
    contentW = ps.PageWidth - ps.LeftMargin - ps.RightMargin
    contentH = ps.PageHeight - ps.TopMargin - ps.BottomMargin

    #maxW
    maxW = contentW * opt['imgRatioWidthMax']

    #maxH, 預設版心高之0.937043054427296倍, 為圖名預留空間; 有給預留格數且模板有行格線時改以格數計算
    maxH = contentH * 0.937043054427296
    if opt['imgHeightReserveLines'] is not None:
        h = getImgHeightMaxByGrid(ps, grid, contentH, opt['imgHeightReserveLines'])
        if h is not None:
            maxH = h
    maxH = maxH * opt['imgRatioHeightMax']

    return maxW, maxH


class LinkedPictureError(Exception):
    #斷開連結後仍為連結圖, 即Word無法解析其來源路徑
    pass


def getLinkedPictures(doc):
    #取出主文全部連結圖之[(LinkFormat, 來源SourceFullName)], 含行內圖片(InlineShapes)與浮動圖片(Shapes)
    #note: 非連結圖取s.LinkFormat不會拋例外, 須讀其SourceFullName才能判斷(非連結圖會拋例外), 故以能否讀到SourceFullName為準
    rs = []
    for shapes in (doc.InlineShapes, doc.Shapes):
        for s in shapes:
            try:
                lf = s.LinkFormat
                rs.append((lf, str(lf.SourceFullName)))
            except:
                pass
    return rs


def embedLinkedPictures(doc):
    #本機圖片內嵌: Word匯入html時本機圖片一律成連結圖(LinkFormat.SavePictureWithDocument=False), 另存docx照存為外部連結(TargetMode="External", a:blip只有r:link),
    #  docx移至他處或刪除原圖後即成空框; 故複製進模板後逐一改為存入文件並斷開連結, data URI之圖片本即內嵌不受影響
    #why: Word只解析反斜線絕對路徑, 相對路徑原字, file:///原字(實測Word 16); 正斜線絕對路徑, 百分比編碼之路徑或網址無法解析, 成24×24之破圖連結, 斷開後仍為連結圖,
    #  此類不得悄悄存為破圖, 回傳其SourceFullName供呼叫端失敗並列出
    for lf, fp in getLinkedPictures(doc):
        try:
            lf.SavePictureWithDocument = True
            lf.BreakLink()
        except:
            pass #失敗者於下面複查仍為連結圖, 一併列為未內嵌

    return [fp for lf, fp in getLinkedPictures(doc)]


def htmlToDocxCore(app, docs, fpInSrc, fpInTemp, fpOut, opt):

    #Open, 開啟之文件推入docs, 供htmlToDocx於finally關閉
    docInTemp = app.Documents.Open(fpInTemp)
    docs.append(docInTemp)
    docInSrc = app.Documents.Open(fpInSrc)
    docs.append(docInSrc)

    # 取得來源文件全文 Range（含格式）
    rngSrc = docInSrc.Range()

    # 取得模板文件的主內容 Range
    rngDst = docInTemp.Content  # 或 docInTemp.Range()

    # 如果你想插在文件最後，可先收合到結尾
    WdCollapseEnd = 0 #wdCollapseEnd, 收合到Range結尾
    rngDst.Collapse(WdCollapseEnd)

    # 直接複製格式化文字（等同貼上，但不經過剪貼簿也不靠 Selection）
    rngDst.FormattedText = rngSrc.FormattedText

    #本機圖片內嵌, 有無法內嵌者即失敗(尚未另存, 不產生輸出檔), 逐一列出來源至stderr
    fpsUnresolved = embedLinkedPictures(docInTemp)
    if len(fpsUnresolved) > 0:
        for fp in fpsUnresolved:
            print(f'unresolved linked picture: {fp}', file=sys.stderr)
        raise LinkedPictureError(f'{len(fpsUnresolved)} linked picture(s) can not be embedded since Word can not resolve their source(see stderr)')

    # #選擇src並複製全文
    # docInSrc.Activate()
    # app.Selection.WholeStory()
    # app.Selection.Copy()

    # #選擇模板並貼上全文
    # docInTemp.Activate()
    # app.Selection.Paste()
    # # WdFormatOriginalFormatting = 16
    # # app.Selection.PasteAndFormat(WdFormatOriginalFormatting)

    #選擇模板調整字型
    try:

        fontFamilies=opt['fontFamilies']

        if isinstance(fontFamilies, (list, tuple)) and len(fontFamilies) > 0:
            #print(fontFamilies)

            rng = docInTemp.Range()
            for fn in fontFamilies:
                # print(fn)
                rng.Font.Name = fn

    except:
        err=getError()
        print(err, file=sys.stderr)

    #選擇模板偵測各圖片寬高是否大於最大值(見getImgSizeMax), 並等比例縮小
    try:

        #grids, 模板各節之文件格線, 僅有給預留格數時需要; 節數與文件不符(轉入內容另有分節)時不採用, 沿用預設算法
        grids = getSectGrids(fpInTemp) if opt['imgHeightReserveLines'] is not None else []
        if len(grids) != docInTemp.Sections.Count:
            grids = []

        #kpMax, 各節之最大寬高
        #why: 以圖片所在節之版面計算, 各節版面不同之模板以docInTemp.PageSetup取值會得wdUndefined
        kpMax = {}

        # 使用 docInTemp.InlineShapes 直接取得整份文件所有圖片
        for s in docInTemp.InlineShapes:
            try:

                #maxW, maxH
                sect = s.Range.Sections(1)
                k = sect.Index
                if k not in kpMax:
                    grid = grids[k - 1] if len(grids) > 0 else None
                    kpMax[k] = getImgSizeMax(sect.PageSetup, grid, opt)
                maxW, maxH = kpMax[k]

                #w, h
                w = s.Width
                h = s.Height

                #scale, 計算縮放比例
                scaleW = maxW / w if w > maxW else 1
                scaleH = maxH / h if h > maxH else 1
                scale = min(scaleW, scaleH)

                if scale < 1:
                    s.LockAspectRatio = True #LockAspectRatio=true時只設一邊, 另一邊依比例調整
                    if scaleH < scaleW:
                        s.Height = h * scale #高度受限時直接設高度, 避免由寬度換算之進位使高度略超上限(行格線下超過即多佔1格)
                    else:
                        s.Width = w * scale

            except:
                pass

    except:
        err=getError()
        print(err, file=sys.stderr)

    #設定前後行距
    try:

        for p in docInTemp.Paragraphs:

            try:
                lf = p.Range.ListFormat
            except:
                #若可能沒有ListFormat須跳過
                continue

            listType = lf.ListType   # 0,1,2,3,4
            listStr  = lf.ListString # '•', '1.', ''...

            # #測試顯示ListType
            # try:
            #     rng = p.Range.Duplicate
            #     rng.End = rng.End - 1 #最末是換行符號, 不能插在換行符號後面
            #     rng.InsertAfter(f"[{listType},{listStr}]")
            # except:
            #     err = getError()
            #     print(err)

            WdListNoNumbering = 0       #普通段落
            WdListSingleLevelNumber = 1 #單層編號清單
            WdListMultiLevel = 2        #多層編號清單
            WdListBullet = 3            #項目符號清單
            WdListOutlineNumbering = 4  #大綱編號
            is_list = (listType != WdListNoNumbering) or (listStr != "")

            if is_list:

                fmt = p.Format

                #關閉自動段前段後距離
                try:
                    fmt.SpaceBeforeAuto = False
                    fmt.SpaceAfterAuto = False
                except:
                    pass

                # 設定為 0.5 行
                fmt.LineUnitBefore = 0.5
                fmt.LineUnitAfter = 0.5

    except:
        err=getError()
        print(err, file=sys.stderr)

    #設定本文段落(p)與清單項目(li)之水平對齊, 已為置中者(如表格標題列, 圖名)不更動, 其餘改為左右對齊(兩端對齊), 使右側邊界齊平較為美觀
    try:

        WdAlignParagraphCenter = 1   #置中對齊
        WdAlignParagraphJustify = 3  #左右對齊
        WdOutlineLevelBodyText = 10  #本文層級, 標題(h1-h6)匯入後為大綱層級1-9
        WdWithInTable = 12           #Information之參數, 判斷是否位於表格內

        for p in docInTemp.Paragraphs:
            try:

                rng = p.Range

                #跳過標題段落
                if p.OutlineLevel != WdOutlineLevelBodyText:
                    continue

                #跳過表格內段落
                if rng.Information(WdWithInTable):
                    continue

                #跳過圖片段落
                if rng.InlineShapes.Count > 0:
                    continue

                #跳過置中段落
                if p.Alignment == WdAlignParagraphCenter:
                    continue

                p.Alignment = WdAlignParagraphJustify

            except:
                pass

    except:
        err=getError()
        print(err, file=sys.stderr)

    #移除零寬空格佔位字元(U+200B): 來源為w-md2html於換行標記空div內插入之佔位字元, 用以撐過Word匯入不被當成空段落丟棄
    #於此移除後該段落成為真正的空段落(乾淨段落標記, 顯示編輯標記時亦無可見字元)
    try:

        z = chr(0x200B) #U+200B 零寬空格

        for p in docInTemp.Paragraphs:
            rng = p.Range.Duplicate
            rng.End = rng.End - 1 #排除段落結尾之段落標記, 避免破壞段落結構
            t = rng.Text
            if z in t:
                rng.Text = t.replace(z, '') #移除佔位字元, 段落標記保留 => 乾淨空段落

    except:
        err=getError()
        print(err, file=sys.stderr)

    #SaveAs
    WdSaveFormat = 16 #wdFormatDocumentDefault=16(docx) #https://learn.microsoft.com/zh-tw/office/vba/api/word.wdsaveformat
    docInTemp.SaveAs2(fpOut, WdSaveFormat)


def htmlToDocx(fpInSrc, fpInTemp, fpOut, opt):

    #Dispatch
    app = wc.Dispatch('Word.Application')

    #正式版須隱藏
    app.Visible = False

    #不詢問使用者
    app.DisplayAlerts = False

    #docs, 已開啟之文件, 由htmlToDocxCore推入
    docs = []

    #任一步驟例外時亦須關閉已開啟之文件並結束Word, 否則殘留之WINWORD會鎖住模板與輸出檔, 使後續轉檔失敗
    try:
        htmlToDocxCore(app, docs, fpInSrc, fpInTemp, fpOut, opt)
    finally:

        # Close
        WdDoNotSaveChanges = 0 #wdDoNotSaveChanges
        for doc in reversed(docs):
            try:
                doc.Close(WdDoNotSaveChanges)
            except:
                pass

        #Quit
        WdSaveOptions = 0 #wdDoNotSaveChanges=0(不儲存擱置中變更) #https://learn.microsoft.com/zh-tw/office/vba/api/word.wdsaveoptions
        try:
            app.Quit(WdSaveOptions)
        except:
            pass


def core(b64):
    import math
    state=''

    try:

        #b642str
        s=b642str(b64)

        #j2o
        o=j2o(s)

        #params
        fpInSrc=o['fpInSrc']
        fpInTemp=o['fpInTemp']
        fpOut=o['fpOut']

        #opt
        opt={}
        opt['fontFamilies']=o['fontFamilies']
        opt['imgRatioWidthMax']=o['imgRatioWidthMax']
        opt['imgRatioHeightMax']=o['imgRatioHeightMax']

        #imgHeightReserveLines, 非負整數才採用, 其餘為None(不啟用)
        #note: 須以get取得, 舊版呼叫端不送此鍵, 而轉檔器一律自master下載, 舊版亦會執行新版轉檔器
        reserveLines=o.get('imgHeightReserveLines')
        if isinstance(reserveLines, bool) or not isinstance(reserveLines, (int, float)) or not math.isfinite(reserveLines) or reserveLines < 0 or reserveLines != int(reserveLines):
            reserveLines=None
        else:
            reserveLines=int(reserveLines)
        opt['imgHeightReserveLines']=reserveLines

        #htmlToDocx
        htmlToDocx(fpInSrc, fpInTemp, fpOut, opt)

        state='success'
    except:
        err=getError()
        state='error: '+err['type'].__name__+': '+str(err['message'])

    return state


def run():

    #stdout與stderr改為utf-8: Python寫pipe時預設用系統ANSI字碼頁(中文Windows為cp950), 呼叫端須猜編碼才能讀, 且訊息含cp950無法編碼之字元時print會拋UnicodeEncodeError;
    #PyInstaller打包後不吃PYTHONIOENCODING與PYTHONUTF8, 只能於此設定
    for f in (sys.stdout, sys.stderr):
        try:
            f.reconfigure(encoding='utf-8', errors='backslashreplace')
        except:
            pass

    #由外部程序呼叫或直接給予檔案路徑
    state=''
    argv=sys.argv
    #argv=['','']
    if len(argv)==2:

        #b64
        b64=sys.argv[1]

        #core
        state=core(b64)

    else:
        #print(sys.argv)
        state='error: invalid length of argv'

    #print & flush
    print(state)
    sys.stdout.flush()

    #失敗時以非0離開碼結束, 使呼叫端(如wsemi之execProcess)能直接以離開碼判定成敗, 不必解析輸出
    if state != 'success':
        sys.exit(1)


if True:
    #正式版

    #run
    run()


if False:
    #產生測試輸入b64

    #inp
    inp={
        'fpInSrc':'./ztmp.html',
        'fpInTemp':'./tmp.docx',
        'fpOut':'./ztmp.docx',
        'fontFamilies': ['標楷體','Times New Roman'], #word使用, 因由左往右設定, 故得先設定標楷體再設定Times New Roman
    }
    # print(o2j(inp))

    #str2b64
    b64=str2b64(o2j(inp))
    print(b64)

    #core
    state=core(b64)

    print(state)

