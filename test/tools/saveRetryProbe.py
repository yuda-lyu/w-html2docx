#測試輔助(可自行以 python -I 執行): 載入srcPython/htmlToDocx.py(以模組載入, 不執行run), 以假Word物件驗證另存重試
#  用法: python -I saveRetryProbe.py <htmlToDocx.py之絕對路徑>
#  前提: 已安裝pywin32(建置轉檔器之相依, 見srcPython/說明.txt); 不啟動Word
#  輸出: 一行JSON, 各鍵為檢查項, 值為觀察結果(由WHtml2docx.test.mjs逐項斷言)
import base64
import contextlib
import importlib.util
import io
import json
import sys
import time
from unittest.mock import MagicMock

import pythoncom

SCODE_CONFLICT = -2146823132 #Word「無法儲存或建立此檔案」
SCODE_IN_USE = -2146822935 #Word「檔案正被另一處理程序使用」(占用輸出檔時之回報, 非暫時性衝突之碼)

spec = importlib.util.spec_from_file_location('htmlToDocxUnderTest', sys.argv[1])
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)


def comErr(scode):
    return pythoncom.com_error(-2147352567, '發生例外狀況。', (0, 'Microsoft Word', 'cannot save', 'wdmain11.chm', 0, scode), None)


def makeDoc(errs, always):
    #假文件(其餘成員任意存取皆可): SaveAs2依序拋出errs中之例外, 其後(或errs耗盡)成功; always為True時永遠拋errs[-1]; st['calls']為SaveAs2之呼叫次數
    doc = MagicMock()
    st = {'errs': list(errs), 'calls': 0}

    def saveAs2(fp, fmt):
        st['calls'] += 1
        if always:
            raise st['errs'][-1]
        if len(st['errs']) > 0:
            raise st['errs'].pop(0)

    doc.SaveAs2.side_effect = saveAs2
    return doc, st


def runWith(errs, always, retryCount, retryInterval, rand=lambda: 0.25):
    #直接呼叫saveAs2WithRetry, 回傳{ ret, raised, calls, sleeps }
    doc, st = makeDoc(errs, always)
    sleeps = []
    r = {'ret': None, 'raised': None}
    try:
        r['ret'] = m.saveAs2WithRetry(doc, 'x.docx', 16, retryCount, retryInterval, sleep=lambda s: sleeps.append(round(s, 6)), rand=rand)
    except BaseException as e:
        r['raised'] = type(e).__name__ + ':' + (str(e.excepinfo[5]) if isinstance(e, pythoncom.com_error) else str(e))
    r['calls'] = st['calls']
    r['sleeps'] = sleeps
    return r


def runCore(inp, errs, always):
    #經core(輸入為b64之json, 與轉檔器之命令列相同)跑完整流程, Word以假物件取代; 回傳{ state, calls, stderrRetryLines, sleeps }
    doc, st = makeDoc(errs, always)
    app = MagicMock()
    app.ProtectedViewWindows.Count = 0
    app.Documents.Open.return_value = doc
    sleeps = []
    base = {
        'fpInSrc': 'C:/nonexistent/in.html',
        'fpInTemp': 'C:/nonexistent/tpl.docx',
        'fpOut': 'C:/nonexistent/out.docx',
        'fontFamilies': ['x'],
        'imgRatioWidthMax': 1,
        'imgRatioHeightMax': 1,
    }
    b64 = base64.b64encode(json.dumps({**base, **inp}).encode('utf-8')).decode('ascii')
    bkDispatchEx = m.wc.DispatchEx
    bkSleep = time.sleep
    m.wc.DispatchEx = lambda name: app
    time.sleep = lambda s: sleeps.append(round(s, 6))
    err = io.StringIO()
    try:
        with contextlib.redirect_stderr(err):
            state = m.core(b64)
    finally:
        m.wc.DispatchEx = bkDispatchEx
        time.sleep = bkSleep
    lines = [l for l in err.getvalue().splitlines() if l.startswith('save retry ')]
    return {'state': state, 'calls': st['calls'], 'retryLines': lines, 'sleeps': sleeps, 'quit': app.Quit.call_count}


C = lambda: comErr(SCODE_CONFLICT)
res = {}

#saveAs2WithRetry
res['u_first_ok'] = runWith([], False, 5, 300)
res['u_two_then_ok'] = runWith([C(), C()], False, 5, 300, rand=lambda: 0.25)
res['u_exhausted'] = runWith([C()], True, 3, 300, rand=lambda: 0.5)
res['u_other_code'] = runWith([comErr(SCODE_IN_USE)], True, 5, 300)
res['u_not_com_error'] = runWith([ValueError('boom')], True, 5, 300)
res['u_count0'] = runWith([C()], True, 0, 300)
res['u_rand_lo'] = runWith([C()], False, 5, 300, rand=lambda: 0.0)
res['u_rand_hi'] = runWith([C()], False, 5, 300, rand=lambda: 0.999999)
res['u_conflict_then_other'] = runWith([C(), comErr(SCODE_IN_USE)], False, 5, 300) #先衝突(重試), 重試時改回非暫時性之碼: 即刻拋出, 不再重試

#core: 轉檔器入口之設定解析與完整流程
res['c_retry_ok'] = runCore({'saveRetryCount': 2, 'saveRetryInterval': 0}, [C(), C()], False)
res['c_exhausted'] = runCore({'saveRetryCount': 2, 'saveRetryInterval': 0}, [C()], True)
res['c_count0'] = runCore({'saveRetryCount': 0, 'saveRetryInterval': 0}, [C()], True)
res['c_default'] = runCore({}, [C()], True) #舊版呼叫端不送此二鍵: 預設5次, 基準300毫秒
res['c_null'] = runCore({'saveRetryCount': None, 'saveRetryInterval': None}, [C()], True)
res['c_interval_custom'] = runCore({'saveRetryCount': 1, 'saveRetryInterval': 1000}, [C()], True)
for k, v in (('neg', -1), ('frac', 1.5), ('str', '3'), ('bool', True), ('inf', float('inf'))):
    res['c_bad_count_' + k] = runCore({'saveRetryCount': v, 'saveRetryInterval': 0}, [C()], True)
for k, v in (('neg', -5), ('str', 'x'), ('bool', True), ('inf', float('inf')), ('big', 2147483648)): #big: 超出計時上限2147483647
    res['c_bad_interval_' + k] = runCore({'saveRetryCount': 1, 'saveRetryInterval': v}, [C()], True)
res['c_interval_max'] = runCore({'saveRetryCount': 1, 'saveRetryInterval': 2147483647}, [C()], True) #計時上限本身有效
res['c_other_code'] = runCore({'saveRetryCount': 5, 'saveRetryInterval': 0}, [comErr(SCODE_IN_USE)], True)

print(json.dumps(res)) #ASCII輸出(含\uXXXX跳脫), 不受管線字碼頁影響
