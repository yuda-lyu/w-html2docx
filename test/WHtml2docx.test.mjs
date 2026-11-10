import fs from 'fs'
import http from 'http'
import path from 'path'
import { spawnSync } from 'child_process'
import w from 'wsemi'
import assert from 'assert'
import { unzipSync, strFromU8 } from 'fflate'
import WHtml2docx from '../src/WHtml2docx.mjs'
import ensureBins from '../src/ensureBins.mjs'
import { setBinEnv } from './tools/binEnv.mjs'
import { genTallPng, genGridTemplate, genFigureHtml, getImgHeights, getFigurePages } from './tools/imgHeightFixture.mjs'
import { LINK_CASES, genLinkCases, genHtml, genLinkHtml, getDocxPictures } from './tools/imgLinkFixture.mjs'
import { sleep, wordPids, isAutomationWord, waitWordGone, killNewAutomationWords, clearDisabledItems, startHeldWord, startBorrower, startVisibleWordWatch, startHoldServer } from './tools/wordInstance.mjs'


function isWindows() {
    return process.platform === 'win32'
}


function getDocumentXml(fp) {
    //docx為zip, 取出主文件word/document.xml
    let u8 = fs.readFileSync(fp)
    let files = unzipSync(new Uint8Array(u8))
    return strFromU8(files['word/document.xml'])
}


function getParagraphs(xml) {
    //切出全部段落<w:p>, 並標記其分類與水平對齊

    //tbls, 各表格之區間
    let tbls = []
    let rTbl = /<w:tbl>[\s\S]*?<\/w:tbl>/g
    let mt
    while ((mt = rTbl.exec(xml)) !== null) {
        tbls.push([mt.index, mt.index + mt[0].length])
    }
    let inTable = (i) => {
        return tbls.some((t) => i >= t[0] && i < t[1])
    }

    //ps, 各段落
    let ps = []
    let rP = /<w:p\b[^>]*\/>|<w:p\b[^>]*>[\s\S]*?<\/w:p>/g
    let mp
    while ((mp = rP.exec(xml)) !== null) {
        let s = mp[0]
        let mj = s.match(/<w:jc w:val="([^"]+)"\s*\/>/)
        ps.push({
            align: mj ? mj[1] : '', //both為左右對齊, center為置中, 空字串為未設定
            isHeading: /<w:outlineLvl\b/.test(s), //h1-h6匯入後帶有大綱層級
            isList: /<w:numPr>/.test(s), //ul或ol之li
            isImage: /<w:drawing\b|<w:pict\b/.test(s), //含圖片(w:drawing)或水平線hr(w:pict之v:rect)
            inTable: inTable(mp.index), //位於表格內
            text: [...s.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((v) => v[1]).join(''),
        })
    }

    return ps
}


function getTables(xml) {
    //切出全部表格<w:tbl>
    return [...xml.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)].map((v) => v[0])
}


function getCells(tbl) {
    //切出表格內全部儲存格<w:tc>, 並標記其文字與是否設定不換行
    return [...tbl.matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)].map((v) => {
        let s = v[0]
        return {
            noWrap: /<w:noWrap\s*\/>/.test(s), //來源html之white-space:nowrap匯入後為w:tcPr內之w:noWrap
            text: [...s.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((v) => v[1]).join(''),
        }
    })
}


describe('WHtml2docx', function() {

    //check
    if (!isWindows()) {
        return
    }

    let fpIn = `./test/ztmp.html`
    let fpOut = `./test/ztmp.docx`
    let imgRatioWidthMax = 0.5
    let fontFamilies = ['標楷體', 'Times New Roman']
    let opt = {
        imgRatioWidthMax,
        fontFamilies,
    }

    //轉出docx後解析word/document.xml, 各測試針對其內容比對, 不使用檔案大小門檻
    let xml = null
    let ps = null

    //restoreBinEnv, htmlToDocx.exe自本機鏡像(取自repo內bin/)取得, 不連外網
    let restoreBinEnv = null

    before(async function() {
        this.timeout(60000)
        restoreBinEnv = setBinEnv('WHtml2docx-bin')
        await WHtml2docx(fpIn, fpOut, opt)
        xml = getDocumentXml(fpOut)
        ps = getParagraphs(xml)
    })

    after(function() {
        w.fsDeleteFile(fpOut)
        restoreBinEnv()
    })

    it('convert', function() {
        assert.strict.deepEqual(true, ps.length > 0)
        assert.strict.deepEqual(true, xml.includes('計畫報告')) //來源html之h1
        assert.strict.deepEqual(true, xml.includes('資料流程自動化')) //來源html之li
    })

    it('本文段落與清單項目, 除已置中者外, 均為左右對齊', function() {
        let rs = ps.filter((p) => !p.isHeading && !p.isImage && !p.inTable && p.align !== 'center')
        assert.strict.deepEqual(true, rs.length > 0)
        let rsErr = rs.filter((p) => p.align !== 'both')
        assert.strict.deepEqual([], rsErr)
    })

    it('清單項目(li)已納入左右對齊', function() {
        let rs = ps.filter((p) => p.isList && !p.inTable && p.align !== 'center')
        assert.strict.deepEqual(true, rs.length > 0)
        let rsErr = rs.filter((p) => p.align !== 'both')
        assert.strict.deepEqual([], rsErr)
    })

    it('標題段落不套用左右對齊', function() {
        let rs = ps.filter((p) => p.isHeading)
        assert.strict.deepEqual(true, rs.length > 0)
        let rsErr = rs.filter((p) => p.align === 'both')
        assert.strict.deepEqual([], rsErr)
    })

    it('表格內段落不套用左右對齊', function() {
        let rs = ps.filter((p) => p.inTable)
        assert.strict.deepEqual(true, rs.length > 0)
        let rsErr = rs.filter((p) => p.align === 'both')
        assert.strict.deepEqual([], rsErr)
    })

    it('已強制更改字型', function() {
        assert.strict.deepEqual(true, xml.includes(`w:eastAsia="${fontFamilies[0]}"`))
        assert.strict.deepEqual(true, xml.includes(`w:ascii="${fontFamilies[1]}"`))
    })

    it('圖片寬度不超過版心寬度乘以imgRatioWidthMax', function() {

        //twContent, 版心寬度, 單位twip
        let mSz = xml.match(/<w:pgSz w:w="(\d+)"/)
        let mMar = xml.match(/<w:pgMar[^>]*w:right="(\d+)"[^>]*w:left="(\d+)"/)
        assert.strict.deepEqual(true, mSz !== null)
        assert.strict.deepEqual(true, mMar !== null)
        let twContent = parseInt(mSz[1]) - parseInt(mMar[1]) - parseInt(mMar[2])

        //emuMax, 圖片寬度上限, 單位EMU, 1twip=635EMU
        let emuMax = twContent * 635 * imgRatioWidthMax

        //cxs, 各圖片寬度, 單位EMU
        let cxs = [...xml.matchAll(/<wp:extent cx="(\d+)"/g)].map((v) => parseInt(v[1]))
        assert.strict.deepEqual(true, cxs.length > 0)

        let rsErr = cxs.filter((cx) => cx > emuMax)
        assert.strict.deepEqual([], rsErr)
    })

    it('表格內設定white-space:nowrap之儲存格已轉為不換行(w:noWrap)', function() {

        //tbl, 來源html之4.3表格, 以其獨有之'編號'欄位定位
        let tbl = getTables(xml).find((v) => v.includes('編號'))
        assert.strict.deepEqual(true, w.isestr(tbl))

        //cells
        let cells = getCells(tbl)
        assert.strict.deepEqual(true, cells.length > 0)

        //來源html內有設定white-space:nowrap之短欄位, 均須為不換行
        let tsNoWrap = ['編號', '項目', '單位', '狀態', '001', '甲項', '公尺', '完成', '002', '乙項', '公噸', '進行']
        let rsErr = tsNoWrap.filter((t) => {
            let c = cells.find((v) => v.text === t)
            return !c || !c.noWrap
        })
        assert.strict.deepEqual([], rsErr)

    })

    it('表格內未設定white-space:nowrap之儲存格維持可換行', function() {

        //tbl, 來源html之4.3表格
        let tbl = getTables(xml).find((v) => v.includes('編號'))
        let cells = getCells(tbl)

        //長文字欄位未設定white-space:nowrap, 須維持可換行
        let rs = cells.filter((v) => v.text.length > 10)
        assert.strict.deepEqual(true, rs.length > 0)
        let rsErr = rs.filter((v) => v.noWrap)
        assert.strict.deepEqual([], rsErr)

    })

    it('已移除零寬空格佔位字元(U+200B)', function() {
        assert.strict.deepEqual(false, xml.includes(String.fromCharCode(0x200B)))
    })

    it('指定非預設字型時確實套用, 且輸出資料夾不存在時自動建立', async function() {
        this.timeout(60000)

        //fdTmp, 測試專用暫存資料夾, 測前確認不存在, 測後刪除
        let fdTmp = `./test/_tmp/WHtml2docx`
        let fpOut2 = `${fdTmp}/sub/ztmp2.docx`
        w.fsDeleteFolder(fdTmp)
        assert.strict.deepEqual(false, w.fsIsFolder(fdTmp))

        //fontFamilies2, 與預設值不同且來源html未使用之字型, 轉出文件出現該字型即證明opt.fontFamilies確實被送入轉檔器
        let fontFamilies2 = ['新細明體', 'Georgia']
        let r = await WHtml2docx(fpIn, fpOut2, { fontFamilies: fontFamilies2 })
        assert.strict.deepEqual('ok', r)
        let xml2 = getDocumentXml(fpOut2)
        assert.strict.deepEqual(true, xml2.includes(`w:eastAsia="${fontFamilies2[0]}"`))
        assert.strict.deepEqual(true, xml2.includes(`w:ascii="${fontFamilies2[1]}"`))

        w.fsDeleteFolder(fdTmp)
    })

})


describe('WHtml2docx input checks', function() {

    //check
    if (!isWindows()) {
        return
    }

    //以下不需Word即可執行, 皆於轉檔前被擋下

    it('fpInHtml不存在時reject並含路徑', async function() {
        let err = null
        await WHtml2docx('./test/no-such-file.html', './test/_tmp/WHtml2docx/x.docx')
            .catch((e) => {
                err = e
            })
        assert.strict.deepEqual(true, typeof err === 'string' && err.includes('fpInHtml[') && err.includes('does not exist'))
    })

    it('fpInTemp有給但不存在時reject並含路徑', async function() {
        let err = null
        await WHtml2docx('./test/ztmp.html', './test/_tmp/WHtml2docx/x.docx', { fpInTemp: './test/no-such-temp.docx' })
            .catch((e) => {
                err = e
            })
        assert.strict.deepEqual(true, typeof err === 'string' && err.includes('fpInTemp[') && err.includes('does not exist'))
        assert.strict.deepEqual(false, w.fsIsFolder('./test/_tmp/WHtml2docx')) //於檢查階段即reject, 不會建立輸出資料夾
    })

    it('htmlToDocx.exe取不到(離線且無快取無鏡像)時reject字串並含錯誤代碼與檔名, 不啟動轉檔器', async function() {
        //規格T6/E2: 轉檔器經ensureBins取得, 離線缺檔為E_OFFLINE_MISSING; 本函數以字串reject(沿用既有契約)
        let fdTmp = path.resolve('./test/_tmp/WHtml2docx-offline')
        let fpOut2 = `${fdTmp}/out/ztmp.docx`
        let keys = ['W_BIN_MIRROR', 'W_BIN_CACHE', 'W_BIN_OFFLINE']
        let bk = {}
        for (let k of keys) {
            bk[k] = process.env[k]
            delete process.env[k]
        }
        process.env.W_BIN_OFFLINE = '1'
        process.env.W_BIN_CACHE = path.join(fdTmp, 'cache')
        let err = null
        try {
            await WHtml2docx('./test/ztmp.html', fpOut2)
                .catch((e) => {
                    err = e
                })
        }
        finally {
            for (let k of keys) {
                if (bk[k] === undefined) {
                    delete process.env[k]
                }
                else {
                    process.env[k] = bk[k]
                }
            }
        }
        let existOut = fs.existsSync(fpOut2)
        fs.rmSync(fdTmp, { recursive: true, force: true, maxRetries: 10, retryDelay: 300 })
        try {
            fs.rmdirSync(path.resolve('./test/_tmp')) //尚有其他測試之暫存時非空, 略過
        }
        catch (e) {}
        assert.strict.deepEqual(true, typeof err === 'string' && err.includes('[E_OFFLINE_MISSING]') && err.includes('htmlToDocx.exe'), String(err))
        assert.strict.deepEqual(false, existOut) //未啟動轉檔器, 無輸出
    })

})


describe('WHtml2docx imgHeightReserveLines', function() {

    //check
    if (!isWindows()) {
        return
    }

    //fdTmp, 本describe專用暫存資料夾, 於after刪除
    let fdTmp = path.resolve('./test/_tmp/WHtml2docx-imgHeight')
    let fpTempGrid = path.resolve(fdTmp, 'tempGrid.docx')
    let fpHtml = path.resolve(fdTmp, 'figures.html')

    //行格線模板由內建模板衍生(見genGridTemplate): A4且上下邊界72點, 版心高697.9點; 行距18點, 每頁格數floor(697.9÷18)=38
    let pitch = 18
    let nLines = 38
    let hDefault = 697.9 * 0.937043054427296 //未啟用時之圖高上限: 版心高×0.937043054427296×imgRatioHeightMax(1)
    let tol = 0.05 //Word之尺寸以twip(0.05點)量化

    let conv = async (name, opt) => {
        let fpOut = path.resolve(fdTmp, `${name}.docx`)
        let r = await WHtml2docx(fpHtml, fpOut, opt)
        assert.strict.deepEqual('ok', r)
        return fpOut
    }

    //assertHeights, 4張長圖皆恰縮至上限(長圖原高遠超上限, 縮放後即等於上限)
    let assertHeights = (fp, hExpect) => {
        let hs = getImgHeights(fp)
        assert.strict.deepEqual(4, hs.length)
        let rsErr = hs.filter((h) => Math.abs(h - hExpect) > tol)
        assert.strict.deepEqual([], rsErr, `expect ${hExpect}, got ${hs.join(',')}`)
    }

    //restoreBinEnv, htmlToDocx.exe自本機鏡像(取自repo內bin/)取得, 不連外網
    let restoreBinEnv = null

    before(function() {
        restoreBinEnv = setBinEnv('WHtml2docx-imgHeight-bin')
        fs.rmSync(fdTmp, { recursive: true, force: true })
        fs.mkdirSync(fdTmp, { recursive: true })
        genGridTemplate(path.resolve('./src/tmp.docx'), fpTempGrid)
        fs.writeFileSync(fpHtml, genFigureHtml(genTallPng().toString('base64')), 'utf8')
    })

    after(function() {
        restoreBinEnv()
        fs.rmSync(fdTmp, { recursive: true, force: true, maxRetries: 10, retryDelay: 300 })
        try {
            fs.rmdirSync(path.resolve('./test/_tmp')) //尚有其他測試之暫存時非空, 略過
        }
        catch (err) {}
    })

    it('行格線模板給3格: 圖高上限為(每頁格數-3)×行距-0.5點, 且每張圖與其圖名(首行與末行)同頁', async function() {
        this.timeout(180000)
        let fp = await conv('grid-R3', { fpInTemp: fpTempGrid, imgHeightReserveLines: 3 })

        //規格: 圖片最大高度改為(每頁格數－本值)×行距－0.5點
        assertHeights(fp, (nLines - 3) * pitch - 0.5)

        //規格: 整頁高之圖下方仍留有本值格數供間隔段與圖名使用; 間隔段1格加2行圖名恰為3格, 即每頁恰好排滿, 圖與圖名仍同頁
        let ps = getFigurePages(fp)
        assert.strict.deepEqual(4, ps.length)
        let rsErr = ps.filter((p) => p.capStart !== p.img || p.capEnd !== p.img)
        assert.strict.deepEqual([], rsErr)
    })

    it('行格線模板未給: 沿用預設上限(版心高×0.937043054427296)', async function() {
        this.timeout(180000)
        let fp = await conv('grid-none', { fpInTemp: fpTempGrid })

        //規格: 預設null(不啟用), 圖片最大高度為版心高乘以0.937043054427296再乘以imgRatioHeightMax
        assertHeights(fp, hDefault)
    })

    it('無行格線模板(內建模板)給3格: 不生效, 沿用預設上限', async function() {
        this.timeout(180000)
        let fp = await conv('builtin-R3', { imgHeightReserveLines: 3 })

        //規格: 模板無行格線時不生效
        assertHeights(fp, hDefault)
    })

    it('給無效值(負數、小數、非數字)或不小於每頁格數(38)時不生效, 沿用預設上限', async function() {
        this.timeout(300000)
        for (let [name, v] of [['neg', -1], ['float', 1.5], ['str', 'abc'], ['all', nLines]]) {
            let fp = await conv(`grid-${name}`, { fpInTemp: fpTempGrid, imgHeightReserveLines: v })

            //規格: 給無效值、或本值不小於每頁格數時不生效
            assertHeights(fp, hDefault)
        }
    })

    it('imgRatioHeightMax乘於其算出之高度', async function() {
        this.timeout(180000)
        let fp = await conv('grid-R3-ratio', { fpInTemp: fpTempGrid, imgHeightReserveLines: 3, imgRatioHeightMax: 0.5 })

        //規格: 有給imgHeightReserveLines且其生效時, 改為其算出之高度乘以imgRatioHeightMax
        assertHeights(fp, ((nLines - 3) * pitch - 0.5) * 0.5)
    })

    it('舊版呼叫端(轉檔器輸入不含imgHeightReserveLines)照常轉檔, 沿用預設上限', async function() {
        this.timeout(180000)

        //轉檔器自master取得, 舊版套件亦執行新版轉檔器, 故轉檔器須容許輸入缺此鍵; 以ensureBins校驗後之路徑(新版實際執行者)為準
        let fpExe = (await ensureBins({ names: 'htmlToDocx.exe' }))['htmlToDocx.exe']
        let fpOut = path.resolve(fdTmp, 'grid-oldcaller.docx')
        let inp = {
            fpInSrc: fpHtml,
            fpInTemp: fpTempGrid,
            fpOut,
            fontFamilies: ['標楷體', 'Times New Roman'],
            imgRatioWidthMax: 1,
            imgRatioHeightMax: 1,
        }
        let sp = spawnSync(fpExe, [Buffer.from(JSON.stringify(inp), 'utf8').toString('base64')], { windowsHide: true, timeout: 170000 })
        assert.strict.deepEqual(0, sp.status, String(sp.stdout))
        assert.strict.deepEqual('success', String(sp.stdout).trim())
        assertHeights(fpOut, hDefault)
    })

})


describe('WHtml2docx local pictures', function() {

    //check
    if (!isWindows()) {
        return
    }

    //規格: 建議w-html2docx調整.md(2026-10-29, 來源審計md2docx-run01): Word匯入html時本機圖片為連結圖, 轉檔須斷開連結內嵌, 解析不到者報錯
    //  可解析: 反斜線絕對路徑, 相對路徑原字, file:///原字; 不可解析: 正斜線絕對路徑, 百分比編碼之路徑或網址(Word 16實測)

    //fdTmp, 本describe專用暫存資料夾, 其下fx/含圖片夾「圖 資料夾」與網頁夾「頁 面」(路徑含中文與空白), 於after刪除
    let fdTmp = path.resolve('./test/_tmp/WHtml2docx-link')
    let fdOut = path.join(fdTmp, 'out')
    let fx = null

    //本機http伺服器(供Word取遠端圖片): /a.png回傳案A1之圖, 其餘404
    let port = 8938
    let server = null
    let urls = []

    let c = (id) => {
        return fx.cases.find((v) => v.id === id)
    }

    let cs = (ids) => {
        return ids.map(c)
    }

    //fpHtmlOf, 以案組出html
    let fpHtmlOf = (name, cases) => {
        return genLinkHtml(cases, path.join(fx.fdHtml, `${name}.html`))
    }

    //convOk, 轉檔須成功, 回傳docx路徑
    let convOk = async (name, fpHtml) => {
        let fpOut = path.join(fdOut, `${name}.docx`)
        let r = await WHtml2docx(fpHtml, fpOut)
        assert.strict.deepEqual('ok', r)
        return fpOut
    }

    //convErr, 轉檔須失敗(reject字串)且不產生輸出檔, 回傳訊息
    let convErr = async (name, fpHtml) => {
        let fpOut = path.join(fdOut, `${name}.docx`)
        let err = null
        await WHtml2docx(fpHtml, fpOut)
            .catch((e) => {
                err = e
            })
        assert.strict.deepEqual('string', typeof err, String(err))
        assert.strict.deepEqual(false, fs.existsSync(fpOut))
        return err
    }

    //assertEmbedded, docx須自足: 圖片關聯無外部連結, a:blip皆為r:embed並指向word/media內之檔, word/media內有各來源圖片(位元組相同)
    let assertEmbedded = (fp, srcs) => {
        let p = getDocxPictures(fp)

        //規格: 圖片關聯無TargetMode="External"
        assert.strict.deepEqual([], p.imageRels.filter((v) => v.external))

        //規格: <a:blip>為r:embed而非r:link, 且參照之關聯指向word/media內存在之檔
        assert.strict.deepEqual(true, p.blips.length >= srcs.length, `blips ${p.blips.length}`)
        for (let b of p.blips) {
            assert.strict.deepEqual('', b.link)
            let rel = p.imageRels.find((v) => v.id === b.embed)
            assert.strict.deepEqual(true, rel !== undefined, `blip embed[${b.embed}]`)
            assert.strict.deepEqual(true, p.media.some((m) => m.name === rel.target), rel.target)
        }

        //規格: word/media有各來源圖片, 位元組與來源檔(data URI者為其解碼內容)相同
        for (let bytes of srcs) {
            assert.strict.deepEqual(true, p.media.some((m) => m.bytes.equals(bytes)), `media ${p.media.map((m) => `${m.name}:${m.bytes.length}`).join(',')}`)
        }

        return p
    }

    //restoreBinEnv, htmlToDocx.exe自本機鏡像(取自repo內bin/)取得, 不連外網
    let restoreBinEnv = null

    before(async function() {
        restoreBinEnv = setBinEnv('WHtml2docx-link-bin')
        fs.rmSync(fdTmp, { recursive: true, force: true })
        fs.mkdirSync(fdOut, { recursive: true })
        fx = genLinkCases(path.join(fdTmp, 'fx'))
        server = http.createServer((req, res) => {
            urls.push(req.url)
            if (req.url === '/a.png') {
                res.writeHead(200, { 'content-type': 'image/png', 'content-length': String(c('A1').bytes.length) })
                res.end(c('A1').bytes)
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
        restoreBinEnv()
        fs.rmSync(fdTmp, { recursive: true, force: true, maxRetries: 10, retryDelay: 300 })
        try {
            fs.rmdirSync(path.resolve('./test/_tmp')) //尚有其他測試之暫存時非空, 略過
        }
        catch (err) {}
    })

    //assertSvgEmbedded, SVG以向量(svg)與點陣預覽(png)內嵌; Word會重寫svg內容(如色碼轉大寫), 故不比對位元組, 只確認為來源之svg(含來源之紅色矩形)
    let assertSvgEmbedded = (p) => {
        let svgs = p.media.filter((m) => m.name.endsWith('.svg')).map((m) => m.bytes.toString('utf8').toLowerCase())
        assert.strict.deepEqual(true, svgs.some((s) => s.includes('<svg') && s.includes('#dc2626')), JSON.stringify(svgs))
        assert.strict.deepEqual(true, p.media.some((m) => m.name.endsWith('.png')))
    }

    for (let id of ['A1', 'R1', 'F2', 'J1', 'S1']) {
        it(`${id} Word可解析之本機路徑寫法(${LINK_CASES.find((v) => v.id === id).desc}): 圖片斷開連結內嵌, docx自足`, async function() {
            this.timeout(180000)
            let fp = await convOk(id, fpHtmlOf(id, cs([id])))
            if (id === 'S1') {
                assertSvgEmbedded(assertEmbedded(fp, []))
            }
            else {
                assertEmbedded(fp, [c(id).bytes])
            }
        })
    }

    for (let id of ['B', 'D2']) {
        it(`${id} ${LINK_CASES.find((v) => v.id === id).desc}: 照舊內嵌, 圖數與顯示尺寸與改前相同`, async function() {
            this.timeout(180000)
            let fp = await convOk(id, fpHtmlOf(id, cs([id])))
            let p = assertEmbedded(fp, [c(id).bytes])

            //規格: data URI之圖片轉檔結果與改前相同; 改前(2026-10-29以現行bin/htmlToDocx.exe轉同夾具)之顯示尺寸: 320×180 px之PNG為240×135點, 240×120 px之JPEG為180×90點
            assert.strict.deepEqual(1, p.extents.length)
            let ex = id === 'B' ? { cx: 240, cy: 135 } : { cx: 180, cy: 90 }
            assert.strict.deepEqual(true, Math.abs(p.extents[0].cx - ex.cx) <= 0.05 && Math.abs(p.extents[0].cy - ex.cy) <= 0.05, JSON.stringify(p.extents))
        })
    }

    for (let [id, mk] of [['A2', 'A2.png'], ['R2', 'R2.png'], ['F1', 'F1.png']]) {
        it(`${id} Word無法解析之寫法(${LINK_CASES.find((v) => v.id === id).desc}): reject並列出該圖來源, 不產生輸出檔`, async function() {
            this.timeout(180000)
            let err = await convErr(id, fpHtmlOf(id, cs([id])))

            //規格: 斷開後仍為連結圖者以非0離開碼結束, 訊息列出SourceFullName(Word解析失敗時為其改寫之路徑, 仍含檔名)
            assert.strict.deepEqual(true, err.startsWith('code[1]:'), err)
            assert.strict.deepEqual(true, err.includes(`unresolved linked picture: `) && err.includes(mk), err)
            assert.strict.deepEqual(true, /\b1 linked picture\(s\)/.test(err), err)
        })
    }

    it('同一份html含10案: reject並只列出3個無法解析者, Word可解析者與data URI不列入', async function() {
        this.timeout(180000)
        let err = await convErr('all10', fpHtmlOf('all10', fx.cases))

        //規格: 列出各無法解析圖片之SourceFullName, 其餘不列
        assert.strict.deepEqual(true, /\b3 linked picture\(s\)/.test(err), err)
        for (let mk of ['A2.png', 'R2.png', 'F1.png']) {
            assert.strict.deepEqual(true, err.includes(mk), `${mk}: ${err}`)
        }
        for (let mk of ['A1.png', 'R1.png', 'F2.png', 'J1.jpg', 'S1.svg', 'B.png', 'D2.jpg']) {
            assert.strict.deepEqual(false, err.includes(mk), `${mk}: ${err}`)
        }
        assert.strict.deepEqual(3, err.split('unresolved linked picture: ').length - 1, err)
    })

    it('同一份html含全部可解析者(5案)與data URI(2案): 7張圖皆內嵌', async function() {
        this.timeout(180000)
        let ids = ['A1', 'R1', 'F2', 'J1', 'S1', 'B', 'D2']
        let fp = await convOk('embed7', fpHtmlOf('embed7', cs(ids)))
        assertSvgEmbedded(assertEmbedded(fp, cs(ids.filter((v) => v !== 'S1')).map((v) => v.bytes)))
    })

    it('圖檔不存在: reject並列出該路徑, 不產生輸出檔', async function() {
        this.timeout(180000)
        let fpNo = path.join(path.dirname(c('A1').fp), '不 存在.png')
        let fpHtml = genHtml(`<p>missing</p><p><img src="${fpNo}"></p>`, path.join(fx.fdHtml, 'missing.html'))
        let err = await convErr('missing', fpHtml)

        //規格: 解析不到之路徑報錯並列出SourceFullName
        assert.strict.deepEqual(true, err.includes(`unresolved linked picture: ${fpNo}`), err)
    })

    it('浮動圖片(align="right"之Shapes)亦內嵌; 其不存在時亦reject並列出路徑', async function() {
        this.timeout(180000)

        //規格: 對全部InlineShapes與Shapes之連結圖斷開連結; Word匯入html時<img align="left|right">為浮動圖片(Shapes, 實測),
        //  style="float:left|right"與position:absolute則仍為行內圖片(InlineShapes), 故以align測Shapes
        let txt = 'text text text text text text text text text text text text'
        let fpHtml = genHtml(`<p>float</p><p><img src="${c('A1').fp}" align="right" width="100"> ${txt}</p>`, path.join(fx.fdHtml, 'float.html'))
        let fp = await convOk('float', fpHtml)
        assertEmbedded(fp, [c('A1').bytes])
        assert.strict.deepEqual(true, getDocumentXml(fp).includes('<wp:anchor'), 'the picture is expected to be a floating shape(wp:anchor)')

        let fpNo = path.join(path.dirname(c('A1').fp), '浮 動不存在.png')
        let fpHtml2 = genHtml(`<p>float</p><p><img src="${fpNo}" align="right" width="100"> ${txt}</p>`, path.join(fx.fdHtml, 'float-missing.html'))
        let err = await convErr('float-missing', fpHtml2)
        assert.strict.deepEqual(true, err.includes(`unresolved linked picture: ${fpNo}`), err)
    })

    it('http圖片: Word取回後內嵌; 取不到(404)時reject並列出網址', async function() {
        this.timeout(180000)

        //規格: 斷開連結時Word自網址取回圖片內嵌, 取不到者仍為連結圖而報錯
        let fpHtml = genHtml(`<p>http</p><p><img src="http://127.0.0.1:${port}/a.png"></p>`, path.join(fx.fdHtml, 'http.html'))
        let fp = await convOk('http', fpHtml)
        assertEmbedded(fp, [c('A1').bytes])
        assert.strict.deepEqual(true, urls.includes('/a.png'), JSON.stringify(urls))

        let fpHtml2 = genHtml(`<p>http</p><p><img src="http://127.0.0.1:${port}/nope.png"></p>`, path.join(fx.fdHtml, 'http-404.html'))
        let err = await convErr('http-404', fpHtml2)
        assert.strict.deepEqual(true, err.includes(`unresolved linked picture: http://127.0.0.1:${port}/nope.png`), err)
    })

})


describe('WHtml2docx Word instance isolation', function() {

    //check
    if (!isWindows()) {
        return
    }

    //規格: 主控排程第408項(hd408, 2026-11-10, 源自下游w-md2docx之調查): 轉檔器每次另開獨立之Word實例(DispatchEx), 只結束本行程所開者
    //  ①他行程已開之Word(含其未存內容)不被借用, 不被結束; ②兩行程並行轉檔(同時, 錯開1.5秒, 交錯)皆成功;
    //  ③例外路徑(轉檔失敗, 來源或模板開啟失敗)亦只結束自己之實例且不殘留WINWORD; ④本行程之實例被他行程(舊寫法Dispatch)連上並開文件時, 轉完不結束該實例
    //  他行程之Word以PowerShell經COM扮演(見tools/wordInstance.mjs), 一律隱藏; 本組須在無其他Word時執行(見before)

    let fdTmp = path.resolve('./test/_tmp/WHtml2docx-word')
    let fdOut = path.join(fdTmp, 'out')

    //png, 1×1之PNG, 由扣住點伺服器回傳
    let png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64')

    //pre, 本組開始時已存在之WINWORD(須為空, 見before); null為本組未執行(略過或初始化前失敗), 此時after不結束任何WINWORD
    let pre = null

    //restoreBinEnv, htmlToDocx.exe自本機鏡像(取自repo內bin/)取得, 不連外網
    let restoreBinEnv = null

    //watch, 可見Word視窗之監看(規格: 不得開啟可見之Word視窗)
    let watch = null

    let fpStopOf = (name) => {
        return path.join(fdTmp, `${name}.stop`)
    }

    let genPage = (name, body) => {
        let fp = path.join(fdTmp, `${name}.html`)
        fs.writeFileSync(fp, `<html><head><meta charset="utf-8"></head><body>${body}</body></html>`, 'utf8')
        return fp
    }

    //genHoldHtml, 含一張自扣住點伺服器取得之圖片: Word匯入時向伺服器取圖, 伺服器扣住期間即為「轉檔進行中」之確定時點
    let genHoldHtml = (name, url) => {
        return genPage(name, `<h1>Word instance ${name}</h1><p>hold</p><p><img src="${url}"></p>`)
    }

    //conv, 轉檔並回傳{ ok, err, fpOut }, 不拋錯
    let conv = async (fpHtml, name) => {
        let fpOut = path.join(fdOut, `${name}.docx`)
        let r = { ok: false, err: '', fpOut }
        await WHtml2docx(fpHtml, fpOut)
            .then((v) => {
                r.ok = v === 'ok'
            })
            .catch((e) => {
                r.err = String(e)
            })
        return r
    }

    //untilHold, 等轉檔進行到Word向伺服器取圖; 轉檔先結束者回傳false
    let untilHold = async (srv, p) => {
        return Promise.race([srv.firstRequest().then(() => true), p.then(() => false)])
    }

    //newPids, 本組開始後新出現之WINWORD
    let newPids = () => {
        return wordPids().filter((v) => !pre.includes(v))
    }

    //assertHeldIntact, H於轉檔後仍在, 其後連續2次讀取皆為文件數1且未存文字不變, 自始無讀取失敗
    let assertHeldIntact = async (h) => {
        let i = h.lines.length
        for (let k = 0; k < 2; k++) {
            let { i: j, line } = await h.waitLine(/^(DOCS|ERR)/, i)
            assert.strict.deepEqual(`DOCS 1 ${h.text}`, line)
            i = j + 1
        }
        assert.strict.deepEqual([], h.lines.filter((v) => /^(ERR|STDERR|SPAWNERR)/.test(v)))
        assert.strict.deepEqual(true, wordPids().includes(h.pid), `held WINWORD ${h.pid}`)
    }

    //assertNoVisible, 規格: 自監看第iw行起無可見之Word視窗(等監看再完成一輪後判定)
    let assertNoVisible = async (iw) => {
        assert.strict.deepEqual([], await watch.visibleSince(iw))
    }

    before(async function() {
        this.timeout(120000)

        //前一組之轉檔器所開之Word可能仍在結束中
        await waitWordGone(wordPids().filter(isAutomationWord), 30000)

        //有其他Word(含使用者開啟者)時不跑本組: 修前(Dispatch)之轉檔器會借用並以不儲存結束之; ④之他行程以GetActiveObject連上時亦無法確定連上本行程之實例
        let pids = wordPids()
        if (pids.length > 0) {
            console.log(`skip [Word instance isolation]: other WINWORD is running(${pids.join(',')})`)
            this.skip()
        }
        pre = pids

        restoreBinEnv = setBinEnv('WHtml2docx-word-bin')

        //並行轉檔前先取得轉檔器至快取, 免兩個轉檔同時寫快取
        await WHtml2docx.ensureBins()

        fs.rmSync(fdTmp, { recursive: true, force: true })
        fs.mkdirSync(fdOut, { recursive: true })

        //先前被強制結束之Word所留下、指向本套件檔案之停用項目, 會使開檔跳出可見訊息框而停住, 先清除
        let rd = clearDisabledItems(path.resolve('.'))
        if (rd.removed.length > 0 || rd.failed.length > 0) {
            console.log(`Word disabled items under this package: ${JSON.stringify(rd)}`)
        }

        watch = await startVisibleWordWatch(fpStopOf('watch'))
    })

    after(async function() {
        this.timeout(60000)
        if (watch) {
            await watch.stop()
        }

        //收尾: 只結束本組新出現且經COM自動化啟動之WINWORD(各案已斷言不殘留, 此處為失敗時之清理); 有強制結束者再清其可能留下之停用項目
        if (pre !== null) {
            let rk = await killNewAutomationWords(pre)
            if (rk.killed.length > 0) {
                let rd = clearDisabledItems(path.resolve('.'))
                console.log(`cleanup: killed WINWORD ${JSON.stringify(rk)}, Word disabled items ${JSON.stringify(rd)}`)
            }
        }
        if (restoreBinEnv) {
            restoreBinEnv()
        }
        fs.rmSync(fdTmp, { recursive: true, force: true, maxRetries: 10, retryDelay: 300 })
        try {
            fs.rmdirSync(path.resolve('./test/_tmp')) //尚有其他測試之暫存時非空, 略過
        }
        catch (err) {}
    })

    it('①他行程已開之Word(含未存文件)不被借用亦不被結束: 轉檔期間另起WINWORD, 轉完原實例仍在且文件數與未存文字不變, 轉檔器之WINWORD不殘留', async function() {
        this.timeout(180000)
        let iw = watch.lines.length
        let h = await startHeldWord(fpStopOf('held1'))
        let srv = await startHoldServer(png)
        try {
            let p = conv(genHoldHtml('held1', srv.url), 'held1')
            assert.strict.deepEqual(true, await untilHold(srv, p), 'conversion ended before Word requested the picture')
            let during = newPids()
            srv.release()
            let r = await p

            //規格①: 轉檔成功, 輸出含來源內容
            assert.strict.deepEqual(true, r.ok, r.err)
            assert.strict.deepEqual(true, getDocumentXml(r.fpOut).includes('Word instance held1'))

            //規格①: 轉檔進行中有H以外之WINWORD, 即轉檔器另起之實例(未借用H)
            let mine = during.filter((v) => v !== h.pid)
            assert.strict.deepEqual(true, during.includes(h.pid) && mine.length >= 1, JSON.stringify({ during, held: h.pid }))

            //規格①: 轉完H仍在, 文件數與未存文字不變
            await assertHeldIntact(h)

            //規格③: 轉檔器所開之WINWORD轉完即結束
            assert.strict.deepEqual([], await waitWordGone(mine))
        }
        finally {
            await srv.close()
            await h.stop()
        }
        assert.strict.deepEqual([], await waitWordGone(newPids()))

        //規格: 不得開啟可見之Word視窗
        await assertNoVisible(iw)
    })

    it('③例外路徑亦只結束自己之實例: 圖片無法解析、來源開啟失敗、模板開啟失敗皆回報錯誤且不產生輸出檔, 他行程之Word仍在且文件不變, 轉檔器之WINWORD不殘留', async function() {
        this.timeout(180000)
        let iw = watch.lines.length
        let h = await startHeldWord(fpStopOf('held3'))
        try {

            //(a)轉檔失敗: 圖檔不存在, 兩份文件皆已開啟後拋LinkedPictureError
            let fpNo = path.join(fdTmp, 'no-such-picture.png')
            let fpHtml = genPage('fail-picture', `<p>fail</p><p><img src="${fpNo}"></p>`)
            let ra = await conv(fpHtml, 'fail-picture')
            assert.strict.deepEqual(false, ra.ok)
            assert.strict.deepEqual(true, ra.err.includes(`unresolved linked picture: ${fpNo}`), ra.err)
            assert.strict.deepEqual(false, fs.existsSync(ra.fpOut))

            //規格③: 每一例外之後他行程之Word皆仍在且文件不變
            await assertHeldIntact(h)

            //(b)來源開啟失敗, (c)模板開啟失敗: 直接呼叫轉檔器(略過WHtml2docx之存在檢查), 轉檔器以非0離開碼回報error, 訊息為Word找不到該檔(含其路徑)
            let prog = (await WHtml2docx.ensureBins())['htmlToDocx.exe']
            let base = {
                fpInSrc: fpHtml,
                fpInTemp: path.resolve('./src/tmp.docx'),
                fpOut: path.join(fdOut, 'fail-open.docx'),
                fontFamilies: ['標楷體', 'Times New Roman'],
                imgRatioWidthMax: 1,
                imgRatioHeightMax: 1,
                imgHeightReserveLines: null,
            }
            for (let [inp, fn] of [[{ ...base, fpInSrc: path.join(fdTmp, 'no-such.html') }, 'no-such.html'], [{ ...base, fpInTemp: path.join(fdTmp, 'no-such-temp.docx') }, 'no-such-temp.docx']]) {
                let rb = spawnSync(prog, [Buffer.from(JSON.stringify(inp), 'utf8').toString('base64')], { encoding: 'utf8', windowsHide: true, timeout: 120000 })
                assert.strict.deepEqual(1, rb.status, `${rb.status} ${rb.stdout} ${rb.stderr}`)
                assert.strict.deepEqual(true, rb.stdout.startsWith('error: ') && rb.stdout.includes(fn), rb.stdout)
                assert.strict.deepEqual(false, fs.existsSync(inp.fpOut))
                await assertHeldIntact(h)
            }

            //規格③: 轉檔器所開之WINWORD皆已結束
            assert.strict.deepEqual([], await waitWordGone(newPids().filter((v) => v !== h.pid)))
        }
        finally {
            await h.stop()
        }
        assert.strict.deepEqual([], await waitWordGone(newPids()))
        await assertNoVisible(iw)
    })

    it('②兩行程交錯轉檔: 甲轉檔進行中(Word正取圖)時乙完整轉完一份, 兩者皆成功且輸出完整, 不殘留WINWORD', async function() {
        this.timeout(180000)
        let iw = watch.lines.length
        let srv = await startHoldServer(png)
        let rA = null
        let rB = null
        let aSettled = false
        try {
            let pA = conv(genHoldHtml('interleaveA', srv.url), 'interleaveA')
            pA.then(() => {
                aSettled = true
            })
            assert.strict.deepEqual(true, await untilHold(srv, pA), 'conversion A ended before Word requested the picture')
            rB = await conv(genPage('interleaveB', '<h1>Word instance interleaveB</h1><p>plain</p>'), 'interleaveB')

            //乙轉完時甲仍被扣住(兩者確實交錯; 乙與甲同用內建模板, 甲之實例已開著它)
            assert.strict.deepEqual(false, aSettled)
            srv.release()
            rA = await pA
        }
        finally {
            await srv.close()
        }

        //規格②: 兩者皆成功, 各自輸出含其來源內容且不含對方之內容
        assert.strict.deepEqual([true, true], [rA.ok, rB.ok], JSON.stringify({ A: rA.err, B: rB.err }))
        let xa = getDocumentXml(rA.fpOut)
        let xb = getDocumentXml(rB.fpOut)
        assert.strict.deepEqual([true, false], [xa.includes('Word instance interleaveA'), xa.includes('interleaveB')])
        assert.strict.deepEqual([true, false], [xb.includes('Word instance interleaveB'), xb.includes('interleaveA')])

        //規格③: 不殘留WINWORD
        assert.strict.deepEqual([], await waitWordGone(newPids()))
        await assertNoVisible(iw)
    })

    it('②兩行程並行轉檔(同時開始、錯開1.5秒, 各2次): 皆成功且輸出完整, 每次結束後不殘留WINWORD', async function() {
        this.timeout(900000)
        let iw = watch.lines.length
        let fpIn = path.resolve('./test/ztmp.html')
        let rs = []
        for (let rep = 0; rep < 2; rep++) {
            for (let off of [0, 1500]) {
                let name = `race-${off}-${rep}`

                //maxWord, 轉檔期間同時存在之轉檔器WINWORD數(取樣觀察, 用以確認兩個實例確有重疊)
                let maxWord = 0
                let smp = { on: true }
                let sampler = (async () => {
                    while (smp.on) {
                        maxWord = Math.max(maxWord, newPids().length)
                        await sleep(300)
                    }
                })()

                //off為情境參數(乙晚甲off毫秒開始), 非同步點
                let [a, b] = await Promise.all([
                    conv(fpIn, `${name}-a`),
                    sleep(off).then(() => conv(fpIn, `${name}-b`)),
                ])
                smp.on = false
                await sampler
                let full = (r) => r.ok && getDocumentXml(r.fpOut).includes('計畫報告') //來源html之h1
                let left = await waitWordGone(newPids())
                rs.push({ name, a: full(a) || a.err, b: full(b) || b.err, maxWord, left })
            }
        }

        //規格②: 0失敗, 兩轉檔各自一個實例且期間重疊(同模板同html); 規格③: 每次結束後不殘留WINWORD
        assert.strict.deepEqual([], rs.filter((v) => v.a !== true || v.b !== true || v.maxWord < 2 || v.left.length > 0), JSON.stringify(rs))
        await assertNoVisible(iw)
    })

    it('④本行程之Word被他行程(舊寫法Dispatch)連上並開文件時: 轉檔成功, 轉完不結束該實例, 對方之文件保留, 對方用畢結束後不殘留', async function() {
        this.timeout(180000)
        let iw = watch.lines.length
        let srv = await startHoldServer(png)
        let b = null
        let x = []
        let r = null
        try {
            let p = conv(genHoldHtml('borrowed', srv.url), 'borrowed')
            assert.strict.deepEqual(true, await untilHold(srv, p), 'conversion ended before Word requested the picture')

            //x, 轉檔器之實例(本組開始時無其他Word, 故此時僅此一個)
            x = newPids()
            assert.strict.deepEqual(1, x.length, JSON.stringify(x))

            //他行程以GetActiveObject連上並新增文件、寫入未存文字(Word取圖期間仍受理COM呼叫)
            b = startBorrower(fpStopOf('borrower'), { text: 'unsaved text of the borrowing client 408' })
            let { line } = await b.waitLine(/^(ADDED|NOATTACH|VISIBLE|EXIT)/)
            assert.strict.deepEqual('ADDED', line, b.lines.join(' / '))
            srv.release()
            r = await p

            //規格④: 轉檔成功, 輸出不含對方文件之文字
            assert.strict.deepEqual(true, r.ok, r.err)
            let xml = getDocumentXml(r.fpOut)
            assert.strict.deepEqual([true, false], [xml.includes('Word instance borrowed'), xml.includes(b.text)])

            //規格④: 對方於轉完後仍持有其同一份文件(文件數1, 以其文件參考讀得其未存文字), 即實例與該文件皆未被轉檔器關閉
            let { line: lt } = await b.waitLine(/^(TEXTSET|SETFAIL|EXIT)/)
            assert.strict.deepEqual('TEXTSET', lt, b.lines.join(' / '))
            await b.stop()
            assert.strict.deepEqual(true, b.lines.includes(`DOCS 1 ${b.text}`), b.lines.join(' / '))
            b = null
        }
        finally {
            if (b) {
                await b.stop()
            }
            await srv.close()
        }

        //對方用畢以Quit結束該實例, 不殘留
        assert.strict.deepEqual([], await waitWordGone(x))
        assert.strict.deepEqual([], await waitWordGone(newPids()))
        await assertNoVisible(iw)
    })

    it('⑤輸出路徑即模板(以輸出覆寫模板)時照常轉檔: 模板改為唯讀開啟之例外, 與改前相同', async function() {
        this.timeout(180000)
        let iw = watch.lines.length

        //規格(不得改變轉檔之輸出與對外API): 改前模板以可寫開啟, fpOut與fpInTemp同檔時覆寫成功; 唯讀開啟則Word拒絕另存回原路徑(實測)
        let fpT = path.join(fdTmp, 'same-path.docx')
        fs.copyFileSync(path.resolve('./src/tmp.docx'), fpT)
        let fpHtml = genPage('same-path', '<h1>Word instance same-path</h1><p>text</p>')
        let r = await WHtml2docx(fpHtml, fpT, { fpInTemp: fpT })
        assert.strict.deepEqual('ok', r)
        assert.strict.deepEqual(true, getDocumentXml(fpT).includes('Word instance same-path'))
        assert.strict.deepEqual([], await waitWordGone(newPids()))
        await assertNoVisible(iw)
    })

})
