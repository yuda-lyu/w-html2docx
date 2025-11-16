from win32com import client as wc  

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
    return base64.b64decode(v)


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


def htmlToDocx(fpInSrc, fpInTemp, fpOut, opt):

    #Dispatch
    app = wc.Dispatch('Word.Application')

    #正式版須隱藏
    app.Visible = False #True False

    #不詢問使用者
    app.DisplayAlerts = False 

    #Open
    docInTemp = app.Documents.Open(fpInTemp)
    docInSrc = app.Documents.Open(fpInSrc)

    #選擇src並複製全文
    docInSrc.Activate()
    app.Selection.WholeStory()
    app.Selection.Copy()

    #選擇模板並貼上全文
    docInTemp.Activate()
    app.Selection.Paste()
    # WdFormatOriginalFormatting = 16
    # app.Selection.PasteAndFormat(WdFormatOriginalFormatting)

    #選擇模板調整字型
    try:
        fontFamilies=opt['fontFamilies']
        if hasattr(fontFamilies, "__len__"):
            if len(fontFamilies)>0:
                #print(fontFamilies)
                docInTemp.Activate()
                app.Selection.WholeStory()
                for fn in fontFamilies:
                    # print(fn)
                    app.Selection.Font.Name = fn
    except:
        err=getError()
        print(err)

    #選擇模板偵測各圖片寬度是否大於滿版(412), 並限制於最大值
    try:
        docInTemp.Activate()
        app.Selection.WholeStory()
        #print(len(app.Selection.Range.InlineShapes)) #有幾張圖
        for s in app.Selection.Range.InlineShapes:
            try:
                s.LockAspectRatio = True
                #240 -> 8.47公分 max=14.55公分 
                if s.Width > 412:
                    s.Width = 412 
            except:
                err = getError()
                print(err)

    except:
        err=getError()
        print(err)

    #SaveAs
    WdSaveFormat = 16 # wdFormatDocumentDefault=16(docx) #https://learn.microsoft.com/zh-tw/office/vba/api/word.wdsaveformat
    docInTemp.SaveAs2(fpOut, WdSaveFormat)

    # Close
    WdDoNotSaveChanges = 0  # wdDoNotSaveChanges
    docInTemp.Close(WdDoNotSaveChanges)
    docInSrc.Close(WdDoNotSaveChanges)

    #Quit
    WdSaveOptions = 0 #wdDoNotSaveChanges=0(不儲存擱置中變更) #https://learn.microsoft.com/zh-tw/office/vba/api/word.wdsaveoptions
    app.Quit(WdSaveOptions)


def core(b64):
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

        #htmlToDocx
        htmlToDocx(fpInSrc, fpInTemp, fpOut, opt)

        state='success'
    except:
        err=getError()
        state='error: '+str(err["message"])

    return state


def run():
    import sys

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

