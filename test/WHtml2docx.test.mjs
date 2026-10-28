import fs from 'fs'
import path from 'path'
import { spawnSync } from 'child_process'
import w from 'wsemi'
import assert from 'assert'
import { unzipSync, strFromU8 } from 'fflate'
import WHtml2docx from '../src/WHtml2docx.mjs'
import ensureBins from '../src/ensureBins.mjs'
import { setBinEnv } from './tools/binEnv.mjs'
import { genTallPng, genGridTemplate, genFigureHtml, getImgHeights, getFigurePages } from './tools/imgHeightFixture.mjs'


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
