import fs from 'fs'
import http from 'http'
import { spawn, execFileSync } from 'child_process'


//Word實例隔離測試之輔助(可移植: 不含本套件之轉檔呼叫, 調用Word之套件可原檔複製使用)
//  驗證對象: 轉檔器須每次另開獨立之Word實例且只結束自己者, 不得借用他行程已開之Word(含使用者開啟者)並於轉完結束之
//  COM用戶端以Windows PowerShell(powershell.exe)扮演: New-Object -ComObject即CoCreateInstance(另起實例, 同DispatchEx), GetActiveObject即舊寫法Dispatch之連上既有實例
//  全部Word一律隱藏(不開可見視窗); 結束時只結束本測試新出現且經COM自動化啟動者(命令列含/Automation或-Embedding)
//  前提: Windows, Windows PowerShell 5.1(powershell.exe), Word 2016以後(登錄機碼Office\16.0), Node 18.2以後(server.closeAllConnections);
//  須在測試閘(Word類, 全機一次一個)內執行: 「新出現之自動化Word」以本組開始時之WINWORD差集判定, 他人於期間另起者無法區分


let sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms))


//wordPids, 目前全部WINWORD.EXE之pid(字串)
function wordPids() {
    let out = execFileSync('tasklist', ['/FI', 'IMAGENAME eq WINWORD.EXE', '/FO', 'CSV', '/NH'], { encoding: 'latin1', windowsHide: true })
    return out.split(/\r?\n/).filter((l) => /^"WINWORD\.EXE"/i.test(l)).map((l) => l.split('","')[1])
}


//isAutomationWord, 是否為經COM自動化啟動之Word(命令列含/Automation或-Embedding); 使用者自行開啟者為否; 查不到(已結束)為否
function isAutomationWord(pid) {
    try {
        let cl = execFileSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', `(Get-CimInstance Win32_Process -Filter "ProcessId=${Number(pid)}").CommandLine`], { encoding: 'utf8', windowsHide: true })
        return /\/Automation|-Embedding/i.test(cl)
    }
    catch (e) {
        return false
    }
}


//waitWordGone, 等待指定pid之WINWORD皆結束, 回傳逾時仍在者
async function waitWordGone(pids, ms = 20000) {
    let t0 = Date.now()
    while (true) {
        let left = wordPids().filter((p) => pids.includes(p))
        if (left.length === 0 || Date.now() - t0 > ms) {
            return left
        }
        await sleep(300)
    }
}


//listDisabledItems, Word之停用項目(HKCU\...\Word\Resiliency\DisabledItems), 回傳[{ name, fp }]; 無此機碼時為[]
//  why: 以taskkill結束開著文件之Word後, Word會將該文件列入停用項目, 此後開啟該檔即跳出可見之「上次開啟時發生嚴重錯誤」訊息框而停住(實測, 模板因此每次轉檔皆停住)
//  值為REG_BINARY: 前12位元組為標頭(第5~8位元組為路徑之位元組數), 其後為UTF-16LE之路徑
let keyDisabledItems = 'HKCU\\Software\\Microsoft\\Office\\16.0\\Word\\Resiliency\\DisabledItems'


function listDisabledItems() {
    let out = ''
    try {
        out = execFileSync('reg', ['query', keyDisabledItems], { encoding: 'latin1', windowsHide: true, stdio: ['ignore', 'pipe', 'ignore'] })
    }
    catch (e) {
        return []
    }
    let rs = []
    for (let l of out.split(/\r?\n/)) {
        let m = l.match(/^\s+(\S+)\s+REG_BINARY\s+([0-9A-Fa-f]+)\s*$/)
        if (!m) {
            continue
        }
        let b = Buffer.from(m[2], 'hex')
        let n = b.length >= 12 ? b.readUInt32LE(4) : 0
        let fp = b.subarray(12, 12 + n).toString('utf16le')
        let i0 = fp.indexOf(String.fromCharCode(0)) //位元組數含結尾之空字元
        rs.push({ name: m[1], fp: i0 >= 0 ? fp.slice(0, i0) : fp })
    }
    return rs
}


//clearDisabledItems, 只移除路徑位於fdRoot之下(本套件之模板與測試夾具)之停用項目, 回傳{ removed, failed }(刪除後重讀確認); 路徑在fdRoot以外者不動
function clearDisabledItems(fdRoot) {
    let root = fdRoot.toLowerCase().replace(/\//g, '\\').replace(/\\+$/, '') + '\\'
    let rs = listDisabledItems().filter((v) => v.fp.toLowerCase().replace(/\//g, '\\').startsWith(root))
    for (let v of rs) {
        try {
            execFileSync('reg', ['delete', keyDisabledItems, '/v', v.name, '/f'], { windowsHide: true, stdio: 'ignore' })
        }
        catch (e) {
            //重讀確認見下
        }
    }
    let names = listDisabledItems().map((v) => v.name)
    return { removed: rs.filter((v) => !names.includes(v.name)), failed: rs.filter((v) => names.includes(v.name)) }
}


//killNewAutomationWords, 收尾: 不在before內且經COM自動化啟動之WINWORD先等其自行結束, 逾時者以taskkill結束並回驗, 回傳{ left, killed, still }
//  以taskkill結束者若開著文件, Word會將其列入停用項目(見listDisabledItems), 呼叫端於其後以clearDisabledItems清除
async function killNewAutomationWords(before, ms = 20000) {
    let news = wordPids().filter((p) => !before.includes(p))
    let left = await waitWordGone(news, ms)
    let killed = []
    for (let pid of left.filter(isAutomationWord)) {
        try {
            execFileSync('taskkill', ['/F', '/T', '/PID', pid], { windowsHide: true, stdio: 'ignore' })
        }
        catch (e) {
            //回驗見下
        }
        killed.push(pid)
    }
    let still = killed.length > 0 ? await waitWordGone(killed, 10000) : []
    return { left, killed, still }
}


//runPsClient, 以PowerShell執行COM用戶端腳本(腳本只用ASCII), 逐行收stdout為事件; stop以旗標檔通知(腳本內輪詢$env:W_STOP); env為額外環境變數(不改process.env)
function runPsClient(script, fpStop, env = {}) {
    fs.rmSync(fpStop, { force: true })
    let p = spawn('powershell.exe', ['-NoProfile', '-NonInteractive', '-EncodedCommand', Buffer.from(script, 'utf16le').toString('base64')], {
        windowsHide: true,
        env: { ...process.env, ...env, W_STOP: fpStop },
        stdio: ['ignore', 'pipe', 'pipe'],
    })
    let lines = []
    let listeners = []
    let buf = ''
    let push = (s) => {
        lines.push(s)
        for (let f of listeners.slice()) {
            f()
        }
    }
    p.stdout.on('data', (d) => {
        buf += String(d)
        let ls = buf.split(/\r?\n/)
        buf = ls.pop()
        for (let l of ls) {
            if (l.trim() !== '') {
                push(l.trim())
            }
        }
    })
    p.stderr.on('data', (d) => {
        push(`STDERR ${String(d).trim().slice(0, 300)}`)
    })
    let exited = false
    let done = new Promise((resolve) => {
        let fin = (code) => {
            if (exited) {
                return
            }
            exited = true
            if (buf.trim() !== '') {
                push(buf.trim())
            }
            push(`EXIT ${code}`)
            resolve(code)
        }
        p.on('close', fin)
        p.on('error', (err) => {
            push(`SPAWNERR ${err.message}`)
            fin(null) //無法啟動時不會有close
        })
    })

    //waitLine, 等第一個(自from行起)符合re之行, 回傳{ i, line }; 逾ms或行程結束仍無則reject(含已收各行)
    let waitLine = (re, from = 0, ms = 60000) => new Promise((resolve, reject) => {
        let timer = null
        let check = () => {
            for (let i = from; i < lines.length; i++) {
                if (re.test(lines[i])) {
                    finish()
                    resolve({ i, line: lines[i] })
                    return
                }
            }
            if (exited) {
                finish()
                reject(new Error(`client exited without ${re}: ${lines.join(' / ')}`))
            }
        }
        let finish = () => {
            clearTimeout(timer)
            listeners = listeners.filter((f) => f !== check)
        }
        timer = setTimeout(() => {
            finish()
            reject(new Error(`timeout waiting for ${re}: ${lines.join(' / ')}`))
        }, ms)
        listeners.push(check)
        check()
    })

    //stop, 通知結束並等待行程結束(逾時則結束其行程樹)
    let stop = async (ms = 30000) => {
        fs.writeFileSync(fpStop, '1')
        let r = await Promise.race([done, sleep(ms).then(() => 'timeout')])
        if (r === 'timeout') {
            try {
                execFileSync('taskkill', ['/F', '/T', '/PID', String(p.pid)], { windowsHide: true, stdio: 'ignore' })
            }
            catch (e) {}
            await done
        }
        fs.rmSync(fpStop, { force: true })
    }

    return { lines, waitLine, stop, done }
}


//startHeldWord, 另起一個隱藏Word(H)並新增一份含未存文字之文件, 代表「他行程已開之Word與其未存內容」
//  事件行: READY(就緒); 其後每250ms一行「DOCS <文件數> <文件文字>」, 讀取失敗為「ERR <訊息>」; stop後關閉文件並結束H, 印END
//  回傳{ pid, text, lines, waitLine, stop }; pid為H之WINWORD(就緒時新出現者)
async function startHeldWord(fpStop, { text = 'unsaved text held by another Word client' } = {}) {
    let before = wordPids()
    let script = [
        `$ProgressPreference = 'SilentlyContinue'`,
        `$w = New-Object -ComObject Word.Application`,
        `$w.Visible = $false`,
        `$w.DisplayAlerts = 0`,
        `$d = $w.Documents.Add()`,
        `$d.Content.Text = $env:W_TEXT`,
        `[Console]::Out.WriteLine('READY')`,
        `for ($i = 0; $i -lt 2400; $i++) {`, //上限10分鐘(防懸置)
        `  Start-Sleep -Milliseconds 250`,
        `  try { [Console]::Out.WriteLine('DOCS ' + $w.Documents.Count + ' ' + $d.Content.Text.Trim()) } catch { [Console]::Out.WriteLine('ERR ' + $_.Exception.Message) }`,
        `  if (Test-Path $env:W_STOP) { break }`,
        `}`,
        `try { $d.Close([ref]0) } catch {}`,
        `try { $w.Quit([ref]0) } catch {}`,
        `[Console]::Out.WriteLine('END')`,
    ].join('\n')
    let c = runPsClient(script, fpStop, { W_TEXT: text })
    await c.waitLine(/^READY$/)
    let pids = wordPids().filter((p) => !before.includes(p))
    return { pid: pids[0], pids, text, ...c }
}


//startBorrower, 以GetActiveObject連上現有Word(舊寫法Dispatch之他行程), 僅限隱藏之Word(可見者為使用者之Word, 不碰), 連上後新增一份文件, 寫入未存文字text並持有
//  事件行: ATTACHED; ADDED(已新增文件); TEXTSET(已寫入text)或SETFAIL <訊息>; NOATTACH(無Word)、VISIBLE(連上者可見, 未動作);
//  stop後以自己持有之文件參考印「DOCS <實例文件數> <該文件之文字>」或「DOCSERR <訊息>」, 關閉文件並Quit該實例(同舊寫法用畢即結束), 印END
function startBorrower(fpStop, { text = 'unsaved text of the borrowing client' } = {}) {
    let script = [
        `$ProgressPreference = 'SilentlyContinue'`,
        `try { $w = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application') } catch { [Console]::Out.WriteLine('NOATTACH ' + $_.Exception.Message); exit 0 }`,
        `[Console]::Out.WriteLine('ATTACHED')`,
        `if ($w.Visible) { [Console]::Out.WriteLine('VISIBLE'); exit 0 }`,
        //Word匯入中(忙碌)時PowerShell之屬性語法取不到型別資訊(實測「找不到屬性 'Text'」), 改以InvokeMember直接經IDispatch讀寫, 並有界重試
        `function GetP($o, $n) { $o.GetType().InvokeMember($n, [Reflection.BindingFlags]::GetProperty, $null, $o, $null) }`,
        `function SetP($o, $n, $v) { [void]$o.GetType().InvokeMember($n, [Reflection.BindingFlags]::SetProperty, $null, $o, @($v)) }`,
        `$d = $w.Documents.Add()`,
        `[Console]::Out.WriteLine('ADDED')`,
        //Word匯入中寫入文字會持續被拒(實測), 故先回報ADDED, 再重試至寫入成功(上限60秒)
        `$ok = $false; $err = ''`,
        `for ($k = 0; ($k -lt 240) -and (-not $ok); $k++) { try { SetP (GetP $d 'Content') 'Text' $env:W_TEXT; $ok = $true } catch { $err = $_.Exception.Message; Start-Sleep -Milliseconds 250 } }`,
        `if ($ok) { [Console]::Out.WriteLine('TEXTSET') } else { [Console]::Out.WriteLine('SETFAIL ' + $err) }`,
        `for ($i = 0; $i -lt 2400; $i++) { Start-Sleep -Milliseconds 250; if (Test-Path $env:W_STOP) { break } }`,
        `try { [Console]::Out.WriteLine('DOCS ' + $w.Documents.Count + ' ' + ([string](GetP (GetP $d 'Content') 'Text')).Trim()) } catch { [Console]::Out.WriteLine('DOCSERR ' + $_.Exception.Message) }`,
        `try { $d.Close([ref]0) } catch {}`,
        `try { $w.Quit([ref]0) } catch {}`,
        `[Console]::Out.WriteLine('END')`,
    ].join('\n')
    return { text, ...runPsClient(script, fpStop, { W_TEXT: text }) }
}


//startVisibleWordWatch, 監看可見之Word視窗: 每200ms列出全部WINWORD之可見頂層視窗, 每個新視窗印一行「VISIBLE <pid>|<類別>|<標題>」,
//  屬自動化實例者(命令列含/Automation或-Embedding)隨即結束之並印「KILLED <pid>」(保護使用者桌面; 其後須以clearDisabledItems清停用項目);
//  每輪印「TICK <n>」供呼叫端等到下一輪(事件同步); 就緒印READY; 短於一輪之閃現可能漏看
async function startVisibleWordWatch(fpStop) {
    let script = [
        `$ProgressPreference = 'SilentlyContinue'`,
        `Add-Type @"`,
        `using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;`,
        `public class WordVisWatch {`,
        `  public delegate bool EnumProc(IntPtr h, IntPtr l);`,
        `  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc f, IntPtr l);`,
        `  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);`,
        `  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);`,
        `  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);`,
        `  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);`,
        `  public static string Esc(string s) { var b = new StringBuilder(); foreach (char ch in s) { if (ch < 32 || ch > 126 || ch == '|') { b.Append("\\\\u" + ((int)ch).ToString("x4")); } else { b.Append(ch); } } return b.ToString(); }`,
        `  public static List<string> List(HashSet<uint> pids) {`,
        `    var r = new List<string>();`,
        `    EnumWindows((h, l) => { uint p; GetWindowThreadProcessId(h, out p); if (pids.Contains(p) && IsWindowVisible(h)) { var c = new StringBuilder(256); GetClassName(h, c, 256); var t = new StringBuilder(256); GetWindowText(h, t, 256); r.Add(h.ToInt64() + "|" + p + "|" + Esc(c.ToString()) + "|" + Esc(t.ToString())); } return true; }, IntPtr.Zero);`,
        `    return r;`,
        `  }`,
        `}`,
        `"@`,
        `$seen = @{}`,
        `[Console]::Out.WriteLine('READY')`,
        `for ($i = 0; $i -lt 18000; $i++) {`, //上限1小時(防懸置)
        `  $ps = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)`,
        `  if ($ps.Count -gt 0) {`,
        `    $set = New-Object 'System.Collections.Generic.HashSet[uint32]'`,
        `    foreach ($x in $ps) { [void]$set.Add([uint32]$x.Id) }`,
        `    foreach ($v in [WordVisWatch]::List($set)) {`,
        `      $f = $v.Split('|')`,
        `      if ($seen.ContainsKey($f[0])) { continue }`,
        `      $seen[$f[0]] = 1`,
        `      [Console]::Out.WriteLine('VISIBLE ' + $f[1] + '|' + $f[2] + '|' + $f[3])`,
        `      $cl = (Get-CimInstance Win32_Process -Filter ('ProcessId=' + $f[1])).CommandLine`,
        `      if ($cl -match '/Automation|-Embedding') { & taskkill.exe /F /T /PID $f[1] | Out-Null; [Console]::Out.WriteLine('KILLED ' + $f[1]) }`,
        `    }`,
        `  }`,
        `  [Console]::Out.WriteLine('TICK ' + $i)`,
        `  if (Test-Path $env:W_STOP) { break }`,
        `  Start-Sleep -Milliseconds 200`,
        `}`,
        `[Console]::Out.WriteLine('END')`,
    ].join('\n')
    let c = runPsClient(script, fpStop)
    await c.waitLine(/^READY$/)

    //visibleSince, 自第i行起之VISIBLE行; 先等其後再一輪(TICK)完成, 使呼叫當下已出現之視窗皆已列入
    let visibleSince = async (i) => {
        let { i: j } = await c.waitLine(/^TICK /, c.lines.length)
        await c.waitLine(/^TICK /, j + 1)
        return c.lines.slice(i).filter((v) => /^(VISIBLE|KILLED)/.test(v))
    }
    return { ...c, visibleSince }
}


//startHoldServer, 本機http伺服器(127.0.0.1, 作業系統配發埠): 對/hold.png之請求一律扣住直到release()才回傳圖片, 供「轉檔進行中」之確定性時點
//  Word匯入html時會向網址取圖, 扣住期間Word仍受理他行程之COM呼叫(實測); firstRequest()於第一個請求到達時resolve
async function startHoldServer(png) {
    let released = false
    let waiters = []
    let urls = []
    let onFirst = null
    let first = new Promise((resolve) => {
        onFirst = resolve
    })
    let server = http.createServer(async (req, res) => {
        urls.push(req.url)
        if (req.url !== '/hold.png') {
            res.writeHead(404)
            res.end()
            return
        }
        onFirst()
        if (!released) {
            await new Promise((resolve) => waiters.push(resolve))
        }
        res.writeHead(200, { 'content-type': 'image/png', 'content-length': String(png.length) })
        res.end(png)
    })
    await new Promise((resolve, reject) => {
        server.once('error', reject)
        server.listen(0, '127.0.0.1', resolve)
    })
    let url = `http://127.0.0.1:${server.address().port}/hold.png`
    let release = () => {
        released = true
        for (let r of waiters.splice(0)) {
            r()
        }
    }
    let close = async () => {
        release()
        server.closeAllConnections()
        await new Promise((resolve) => server.close(resolve))
    }
    return { url, urls, firstRequest: () => first, release, close }
}


export { sleep, wordPids, isAutomationWord, waitWordGone, killNewAutomationWords, listDisabledItems, clearDisabledItems, startHeldWord, startBorrower, startVisibleWordWatch, startHoldServer }
