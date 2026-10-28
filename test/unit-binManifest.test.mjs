import fs from 'fs'
import assert from 'assert'
import { execSync } from 'child_process'
import fsGetBinManifestEntry from 'wsemi/src/fsGetBinManifestEntry.mjs'
import fsEnsureBins from 'wsemi/src/fsEnsureBins.mjs'


//規格: 外部執行檔之存放下載校驗與離線(制式規格) S1、S2、M1、M2、M4、§8 T1、T2
describe(`bin/manifest.json`, function() {
    this.timeout(120000)

    let manifest = JSON.parse(fs.readFileSync('bin/manifest.json', 'utf8'))
    let fdt = './test/_tmp/binManifest'

    after(function() {
        fs.rmSync(fdt, { recursive: true, force: true })
        try {
            fs.rmdirSync('./test/_tmp') //尚有其他測試之暫存時非空, 略過
        }
        catch (err) {}
    })

    it(`should pass the schema check of fsEnsureBins (M1, M2)`, async function() {
        //離線且快取為空資料夾: 清單有效時為E_OFFLINE_MISSING(缺檔), 清單無效時為E_MANIFEST; 不連網
        let err = null
        try {
            await fsEnsureBins('bin/manifest.json', { pkgRoot: '.', cacheDir: `${fdt}/cache`, mirror: '', offline: true })
        }
        catch (e) {
            err = e
        }
        assert.strict.equal(err && err.code, 'E_OFFLINE_MISSING', err && err.message)
    })

    it(`should be a pinned manifest of w-html2docx that lists htmlToDocx.exe (M1, M2, U1)`, function() {
        //M1: pkg、mode、repo; M2/F15: 釘選模式之自有檔網址ref須為40碼提交
        assert.strict.equal(manifest.pkg, 'w-html2docx')
        assert.strict.equal(manifest.mode, 'pinned')
        assert.strict.equal(manifest.repo, 'yuda-lyu/w-html2docx')
        assert.strict.equal(/^[0-9a-f]{40}$/.test(manifest.commit), true, manifest.commit)
        assert.strict.deepStrictEqual(manifest.files.map((v) => [v.name, v.path]), [['htmlToDocx.exe', 'bin/htmlToDocx.exe']])
    })

    it(`should match the real files in bin/ (M4, T2)`, async function() {
        for (let f of manifest.files) {
            for (let p of [f, ...(f.parts || [])]) {
                if (typeof p.path === 'string') {
                    let e = await fsGetBinManifestEntry(p.path)
                    assert.strict.deepStrictEqual([e.sha256, e.size], [p.sha256, p.size], p.path)
                }
            }
        }
    })

    it(`should keep the legacy src/htmlToDocx.exe in the repo for old versions (S2)`, function() {
        //S2舊版相容: 舊版於執行時自主幹src/htmlToDocx.exe下載, 該檔須留在repo(不入npm包, 見下一條)
        assert.strict.equal(fs.statSync('src/htmlToDocx.exe').isFile(), true)
    })

    it(`should not download at install time: no postinstall script (D1)`, function() {
        let pkg = JSON.parse(fs.readFileSync('package.json', 'utf8'))
        assert.strict.equal(Object.prototype.hasOwnProperty.call(pkg.scripts, 'postinstall'), false)
        assert.strict.equal(fs.existsSync('scripts/install.mjs'), false)
    })

    it(`should keep the cache out of git and the npm package (S5) and bin/ out of the npm package except the manifest (S1)`, function() {
        let lines = (fn) => {
            return fs.readFileSync(fn, 'utf8').split(/\r?\n/).map((v) => v.trim())
        }
        assert.strict.equal(lines('.gitignore').includes('/bin-cache/'), true)
        assert.strict.equal(lines('.npmignore').includes('/bin-cache/'), true)
        assert.strict.equal(lines('.npmignore').includes('/test/_tmp/'), true)
        assert.strict.equal(lines('.npmignore').includes('/bin/*'), true)
        assert.strict.equal(lines('.npmignore').includes('!/bin/manifest.json'), true)
        assert.strict.equal(lines('.gitattributes').includes('bin/** binary'), true) //F12
    })

    it(`should not pack external executables into the npm package, only bin/manifest.json (S1, T1)`, function() {
        //測試暫存(鏡像、快取)可能放有執行檔副本(他測試平行執行時即是), 放一個誘餌確保其不入包
        fs.mkdirSync(`${fdt}/leak`, { recursive: true })
        fs.writeFileSync(`${fdt}/leak/htmlToDocx.exe`, 'MZ')
        let r = JSON.parse(execSync('npm pack --dry-run --json', { encoding: 'utf8', windowsHide: true }))
        let info = Array.isArray(r) ? r[0] : Object.values(r)[0] //npm 11以前為陣列, npm 12起為以套件名為鍵之物件
        let fps = info.files.map((v) => v.path)
        assert.strict.deepStrictEqual(fps.filter((v) => /\.(exe|dll|msi|zip|7z|\d{3})$/i.test(v)), [])
        assert.strict.deepStrictEqual(fps.filter((v) => v.startsWith('bin/') || v.startsWith('bin-cache/')), ['bin/manifest.json'])
    })

})
