import fs from 'fs'
import http from 'http'
import path from 'path'
import crypto from 'crypto'
import assert from 'assert'
import ensureBins from '../src/ensureBins.mjs'
import autoDownloadFiles from '../src/autoDownloadFiles.mjs'
import WHtml2docx from '../src/WHtml2docx.mjs'
import { setBinEnv } from './tools/binEnv.mjs'


//規格: 外部執行檔之存放下載校驗與離線(制式規格) D1、D3、D4、O1、O2、O3、E2、E3、§8 T3~T7(T8為主幹模式, 本套件為釘選模式, 不適用)
//以本機HTTP伺服器扮raw.githubusercontent.com(自有檔網址為<repo>/<commit>/<path>), 以本地資料夾扮鏡像, 全程不連外網; 取得之檔為repo內bin/htmlToDocx.exe之實檔
describe(`ensureBins`, function() {
    this.timeout(180000)

    let port = 8937
    let fdt = path.resolve('./test/_tmp/ensureBins')
    let keys = ['W_BIN_MIRROR', 'W_BIN_CACHE', 'W_BIN_OFFLINE']
    let bkEnv = {}

    let manifest = JSON.parse(fs.readFileSync('bin/manifest.json', 'utf8'))
    let ent = manifest.files[0]
    let name = ent.name
    let bytes = fs.readFileSync(ent.path)
    let bad = Buffer.from(bytes) //與實檔同大小而內容不同
    bad[1000] = bad[1000] ^ 0xff
    let urlPath = `/${manifest.repo}/${manifest.commit}/${ent.path}`

    let body = bytes //伺服器目前回傳之內容
    let urls = [] //伺服器收到之請求路徑
    let server = null
    let seq = 0
    let cacheDir = ''

    let sha256Of = (fp) => {
        return crypto.createHash('sha256').update(fs.readFileSync(fp)).digest('hex')
    }

    //fetchLocal, 僅放行raw.githubusercontent.com之網址並改指本機伺服器, 其餘一律拒絕(確保測試不連外網)
    let fetchLocal = (url, init) => {
        let u = String(url)
        if (!u.startsWith('https://raw.githubusercontent.com/')) {
            return Promise.reject(new Error(`unexpected url[${u}]`))
        }
        return globalThis.fetch(u.replace('https://raw.githubusercontent.com', `http://127.0.0.1:${port}`), init)
    }

    //noNet, 被呼叫即代表連網
    let nNet = 0
    let noNet = () => {
        nNet++
        return Promise.reject(new Error('network must not be used'))
    }

    let putMirror = (fdMirror, buf) => {
        let fd = path.join(fdMirror, manifest.pkg, ent.sha256)
        fs.mkdirSync(fd, { recursive: true })
        fs.writeFileSync(path.join(fd, name), buf)
        return path.join(fd, name)
    }

    let listFiles = (fd) => {
        if (!fs.existsSync(fd)) {
            return []
        }
        return fs.readdirSync(fd, { recursive: true, withFileTypes: true }).filter((v) => v.isFile()).map((v) => path.join(v.parentPath, v.name))
    }

    before(async function() {
        fs.rmSync(fdt, { recursive: true, force: true })
        fs.mkdirSync(fdt, { recursive: true })
        server = http.createServer((req, res) => {
            urls.push(req.url)
            if (req.url === urlPath) {
                res.writeHead(200, { 'content-length': String(body.length) })
                res.end(body)
            }
            else {
                res.writeHead(404)
                res.end()
            }
        })
        await new Promise((resolve, reject) => {
            server.once('error', reject)
            server.listen(port, '127.0.0.1', resolve)
        })
    })

    after(async function() {
        await new Promise((resolve) => {
            server.close(resolve)
        })
        fs.rmSync(fdt, { recursive: true, force: true, maxRetries: 10, retryDelay: 300 })
        try {
            fs.rmdirSync(path.resolve('./test/_tmp')) //尚有其他測試之暫存時非空, 略過
        }
        catch (err) {}
    })

    beforeEach(function() {
        for (let k of keys) {
            bkEnv[k] = process.env[k]
            delete process.env[k]
        }
        body = bytes
        urls.length = 0
        nNet = 0
        seq++
        cacheDir = path.join(fdt, `cache${seq}`)
    })

    afterEach(function() {
        for (let k of keys) {
            if (bkEnv[k] === undefined) {
                delete process.env[k]
            }
            else {
                process.env[k] = bkEnv[k]
            }
        }
    })

    it(`T3: first call downloads from the pinned-commit url of GitHub, verifies and caches; second call hits the verified cache without network (D3, D4, O2)`, async function() {
        //規格D3: 快取位置<快取>/<sha256>/<name>; O2: 快取未命中時自GitHub(釘選提交)取得
        let r1 = await ensureBins({ fetch: fetchLocal, cacheDir, mirror: '', offline: false })
        let fp = path.join(cacheDir, ent.sha256, name)
        assert.strict.equal(r1[name], fp)
        assert.strict.equal(sha256Of(fp), ent.sha256)
        assert.strict.equal(fs.statSync(fp).size, ent.size)
        assert.strict.deepStrictEqual(urls, [urlPath]) //M2: 網址ref為清單之40碼提交
        assert.strict.deepStrictEqual(listFiles(cacheDir), [fp]) //只存最終檔, 無暫存殘留

        //規格D4: 第二次不連網(快取命中並校驗)
        let r2 = await ensureBins({ fetch: noNet, cacheDir, mirror: '', offline: false })
        assert.strict.equal(r2[name], fp)
        assert.strict.equal(nNet, 0)
        assert.strict.equal(urls.length, 1)
    })

    it(`T3: concurrent calls share one download (E2)`, async function() {
        //原autoDownloadFiles以pmDownload使併發呼叫共用同一次下載, 改用ensureBins後須維持
        let rs = await Promise.all([1, 2, 3].map(() => ensureBins({ fetch: fetchLocal, cacheDir, mirror: '', offline: false })))
        assert.strict.deepStrictEqual(rs.map((r) => r[name]), [1, 2, 3].map(() => path.join(cacheDir, ent.sha256, name)))
        assert.strict.equal(urls.length, 1)
    })

    it(`T4: content or size not matching the manifest is refused and leaves no file (D2, E3)`, async function() {
        for (let [label, b, code] of [
            ['same size, different content', bad, 'E_HASH_MISMATCH'],
            ['truncated', bytes.subarray(0, bytes.length - 10), 'E_SIZE_MISMATCH'],
            ['longer', Buffer.concat([bytes, Buffer.from('x')]), 'E_SIZE_MISMATCH'],
        ]) {
            body = b
            let err = await ensureBins({ fetch: fetchLocal, cacheDir, mirror: '', offline: false, maxAttempts: 1, retryDelayMs: 0 })
                .then(() => null)
                .catch((e) => e)
            assert.strict.equal(err instanceof Error, true, label)
            assert.strict.equal(err.code, code, `${label}: ${err && err.message}`)
            assert.strict.deepStrictEqual(listFiles(cacheDir), [], label)
        }
    })

    it(`T5: W_BIN_MIRROR (local folder) is used before GitHub, and the mirror copy is verified (O1, O2)`, async function() {
        let fdMirror = path.join(fdt, `mirror${seq}`)
        let fpMirror = putMirror(fdMirror, bytes)
        process.env.W_BIN_MIRROR = fdMirror

        //鏡像命中: 不連網
        let r = await ensureBins({ fetch: noNet, cacheDir, offline: false })
        assert.strict.equal(r[name], path.join(cacheDir, ent.sha256, name))
        assert.strict.equal(sha256Of(r[name]), ent.sha256)
        assert.strict.equal(nNet, 0)
        assert.strict.equal(fs.existsSync(fpMirror), true)
    })

    it(`T5: a tampered mirror copy is refused; online it falls back to GitHub, offline it is E_OFFLINE_MISSING (O2, O3)`, async function() {
        let fdMirror = path.join(fdt, `mirror${seq}`)
        putMirror(fdMirror, bad)
        process.env.W_BIN_MIRROR = fdMirror

        //連網時: 拒用鏡像, 改自GitHub取得正確之檔
        let r = await ensureBins({ fetch: fetchLocal, cacheDir, offline: false })
        assert.strict.equal(sha256Of(r[name]), ent.sha256)
        assert.strict.deepStrictEqual(urls, [urlPath])

        //離線時: 鏡像是唯一來源, 其不符即視為缺檔, 不留檔
        let cacheDir2 = path.join(fdt, `cache${seq}b`)
        let err = await ensureBins({ fetch: noNet, cacheDir: cacheDir2, offline: true })
            .then(() => null)
            .catch((e) => e)
        assert.strict.equal(err && err.code, 'E_OFFLINE_MISSING', err && err.message)
        assert.strict.equal(nNet, 0)
        assert.strict.deepStrictEqual(listFiles(cacheDir2), [])
    })

    it(`T6: W_BIN_OFFLINE=1 without cache or mirror is E_OFFLINE_MISSING naming the file, sha256, size and the path to place it (O3, D3)`, async function() {
        process.env.W_BIN_OFFLINE = '1'
        process.env.W_BIN_CACHE = path.join(fdt, `shared${seq}`)
        let err = await ensureBins({ fetch: noNet })
            .then(() => null)
            .catch((e) => e)
        assert.strict.equal(err instanceof Error, true)
        assert.strict.equal(err.code, 'E_OFFLINE_MISSING', err.message)
        assert.strict.equal(nNet, 0)
        assert.strict.equal(urls.length, 0)

        //D3: 共用快取之結構為<W_BIN_CACHE>/<套件>/<sha256>/<name>
        let m = err.missing[0]
        assert.strict.equal(m.name, name)
        assert.strict.equal(m.sha256, ent.sha256)
        assert.strict.equal(m.size, ent.size)
        assert.strict.equal(m.cachePath, path.join(process.env.W_BIN_CACHE, manifest.pkg, ent.sha256, name))
        for (let s of [name, ent.sha256, String(ent.size), m.cachePath]) {
            assert.strict.equal(err.message.includes(s), true, `message should include [${s}]`)
        }
    })

    it(`T6: with an offline mirror configured, the error also names the mirror path (O3, O1)`, async function() {
        let fdMirror = path.join(fdt, `mirror${seq}`) //不存在之鏡像
        let err = await ensureBins({ fetch: noNet, cacheDir, mirror: fdMirror, offline: true })
            .then(() => null)
            .catch((e) => e)
        assert.strict.equal(err && err.code, 'E_OFFLINE_MISSING', err && err.message)
        let fpMirror = path.join(fdMirror, manifest.pkg, ent.sha256, name)
        assert.strict.equal(err.missing[0].mirrorPath, fpMirror)
        assert.strict.equal(err.message.includes(fpMirror), true)
    })

    it(`T7: a tampered cache file is refused and fetched again (D4)`, async function() {
        let r = await ensureBins({ fetch: fetchLocal, cacheDir, mirror: '', offline: false })
        let fp = r[name]
        assert.strict.equal(urls.length, 1)

        //同大小而內容不同
        fs.writeFileSync(fp, bad)
        fs.utimesSync(fp, 1, 1)
        let r2 = await ensureBins({ fetch: fetchLocal, cacheDir, mirror: '', offline: false })
        assert.strict.equal(r2[name], fp)
        assert.strict.equal(sha256Of(fp), ent.sha256)
        assert.strict.equal(urls.length, 2)

        //被截斷(大小不同)
        fs.writeFileSync(fp, bytes.subarray(0, 100))
        fs.utimesSync(fp, 1, 1)
        let r3 = await ensureBins({ fetch: fetchLocal, cacheDir, mirror: '', offline: false })
        assert.strict.equal(sha256Of(r3[name]), ent.sha256)
        assert.strict.equal(urls.length, 3)
    })

    it(`T7: offline, a tampered cache file is never returned (D4, O3)`, async function() {
        let r = await ensureBins({ fetch: fetchLocal, cacheDir, mirror: '', offline: false })
        fs.writeFileSync(r[name], bad)
        fs.utimesSync(r[name], 1, 1)
        let err = await ensureBins({ fetch: noNet, cacheDir, mirror: '', offline: true })
            .then(() => null)
            .catch((e) => e)
        assert.strict.equal(err && err.code, 'E_OFFLINE_MISSING', err && err.message)
        assert.strict.equal(nNet, 0)
    })

})


//規格: E2(套件端只做提供清單、呼叫fsEnsureBins、匯出ensureBins)、D1(不得於安裝時下載)
describe(`ensureBins wiring in w-html2docx`, function() {
    this.timeout(180000)

    let restore = null

    before(function() {
        restore = setBinEnv('ensureBins-wiring')
    })

    after(function() {
        restore()
    })

    it(`autoDownloadFiles (kept for deep importers) resolves the same verified path as ensureBins (E2)`, async function() {
        let bins = await ensureBins({ names: 'htmlToDocx.exe' })
        let { fpExe } = await autoDownloadFiles()
        assert.strict.equal(fpExe, bins['htmlToDocx.exe'])
        assert.strict.equal(fpExe.startsWith(path.resolve(process.env.W_BIN_CACHE)), true) //取自快取, 非src/
    })

    it(`WHtml2docx exposes ensureBins as a property so the default export (UMD) stays unchanged (E2)`, function() {
        assert.strict.equal(WHtml2docx.ensureBins, ensureBins)
    })

})
