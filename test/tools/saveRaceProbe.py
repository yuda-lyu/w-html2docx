#測試輔助(可自行以 python -I -B 執行): 兩個真實Word實例(各在獨立行程)於同一刻SaveAs2至同一資料夾, 驗證轉檔器之saveAs2WithRetry於真實並行衝突下不失敗
#  用法: python -I -B saveRaceProbe.py <htmlToDocx.py之絕對路徑> <模板docx之絕對路徑> <工作夾之絕對路徑> <輪數> [額外字元數, 預設0]
#  前提: 已安裝pywin32與Word 2016以後; 須在Word測試閘內執行(兩個隱藏之自動化實例, 結束時各自Quit)
#  why: Word之並行另存衝突(兩實例同時另存至同一資料夾, 其一於數十毫秒內回報「無法儲存或建立此檔案」)無法以檔案鎖扮演, 故以兩個真實實例於同刻另存重現;
#    實測(無重試, 每次30輪共60次另存): 模板原樣(額外字元0)每次11~19次失敗, 加2000字8~16次, 加4萬字1~6次, 加40萬字0~1次, 加200萬字0次——文件愈大另存愈久而兩實例愈易錯開, 故用最小者;
#    同一行程內之兩個執行緒因GIL切換(5毫秒)精準度不足, 衝突率明顯較低, 故各用獨立行程, 以約定之epoch時刻自旋等待
#  流程: 父行程起兩個子行程(角色a,b), 各自開Word與文件後印READY並讀stdin之起始epoch; 父行程待兩者READY後送出起始時刻(約1秒後), 其後每0.8秒一輪, 兩子行程於各輪時刻同時SaveAs2
#  輸出(父行程): 一行JSON { saves, retried, failed[], bad[], errs[] }; retried為經重試才成功之次數(0表示本次未發生衝突, 驗證無鑑別力)
import importlib.util
import json
import os
import shutil
import subprocess
import sys
import time
import zipfile

PERIOD = 0.8 #每輪間隔秒數, 須大於一次重試之最長等待(基準300毫秒乘以[1,2)倍率)與另存時間, 使各輪皆能重新對齊


def loadModule(fpPy):
    spec = importlib.util.spec_from_file_location('htmlToDocxUnderTest', fpPy)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


def child(role, fpPy, fpTpl, fd, n, chars):
    import pythoncom
    from win32com import client as wc

    m = loadModule(fpPy)
    res = []
    app = None
    doc = None
    try:
        app = wc.DispatchEx('Word.Application')
        app.Visible = False
        app.DisplayAlerts = False
        fpT = os.path.join(fd, f'tpl-{role}.docx')
        shutil.copyfile(fpTpl, fpT)
        doc = app.Documents.Open(fpT, False, True)
        if chars > 0:
            doc.Content.InsertAfter('x ' * (chars // 2))
        print('READY', flush=True)
        t0 = float(sys.stdin.readline())
        for i in range(n):
            while time.time() < t0 + i * PERIOD:
                pass #自旋至約定時刻
            fp = os.path.join(fd, f'r{i}-{role}.docx')
            try:
                k = m.saveAs2WithRetry(doc, fp, 16, m.SAVE_RETRY_COUNT_DEFAULT, m.SAVE_RETRY_INTERVAL_DEFAULT)
                res.append({'i': i, 'retries': k, 'fp': fp})
            except BaseException as e:
                res.append({'i': i, 'error': repr(e)[:200], 'fp': fp})
    except BaseException as e:
        res.append({'i': -1, 'err': repr(e)[:300], 'fp': ''})
    finally:
        for f in (lambda: doc.Close(0), lambda: app.Quit(0)):
            try:
                f()
            except BaseException:
                pass
    print('RESULT ' + json.dumps(res), flush=True)


def isDocx(fp):
    try:
        with zipfile.ZipFile(fp) as z:
            return b'</w:document>' in z.read('word/document.xml')
    except BaseException:
        return False


def parent(fpPy, fpTpl, fd, n, chars):
    os.makedirs(fd, exist_ok=True)
    ps = {}
    for role in ('a', 'b'):
        ps[role] = subprocess.Popen([sys.executable, '-I', '-B', os.path.abspath(__file__), '--child', role, fpPy, fpTpl, fd, str(n), str(chars)], stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, creationflags=subprocess.CREATE_NO_WINDOW)

    #待兩者READY(開Word與文件需數秒, 逾時視為失敗)
    errs = []
    for role, p in ps.items():
        line = p.stdout.readline().strip()
        if line != 'READY':
            errs.append(f'{role}: not ready: {line!r}')
    if len(errs) == 0:
        t0 = time.time() + 1.0
        for p in ps.values():
            p.stdin.write(f'{t0}\n')
            p.stdin.flush()

    saves = []
    for role, p in ps.items():
        out, err = p.communicate(timeout=300)
        rs = [l for l in out.splitlines() if l.startswith('RESULT ')]
        if len(rs) == 0:
            errs.append(f'{role}: no result: {err.strip()[:300]}')
            continue
        for v in json.loads(rs[-1][7:]):
            if v.get('i') == -1:
                errs.append(f'{role}: {v.get("err")}')
            else:
                saves.append(v)

    failed = [v['error'] for v in saves if 'error' in v]
    retried = len([v for v in saves if v.get('retries', 0) > 0])
    bad = [os.path.basename(v['fp']) for v in saves if 'error' not in v and not isDocx(v['fp'])]
    print(json.dumps({'saves': len(saves), 'retried': retried, 'failed': failed, 'bad': bad, 'errs': errs}))


if sys.argv[1] == '--child':
    child(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], int(sys.argv[6]), int(sys.argv[7]))
else:
    parent(sys.argv[1], sys.argv[2], sys.argv[3], int(sys.argv[4]), int(sys.argv[5]) if len(sys.argv) > 5 else 0)
