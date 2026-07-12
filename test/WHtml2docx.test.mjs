import fs from 'fs'
import w from 'wsemi'
import assert from 'assert'
import { unzipSync, strFromU8 } from 'fflate'
import WHtml2docx from '../src/WHtml2docx.mjs'


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

    before(async function() {
        this.timeout(60000)
        await WHtml2docx(fpIn, fpOut, opt)
        xml = getDocumentXml(fpOut)
        ps = getParagraphs(xml)
    })

    after(function() {
        w.fsDeleteFile(fpOut)
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

    it('已移除零寬空格佔位字元(U+200B)', function() {
        assert.strict.deepEqual(false, xml.includes(String.fromCharCode(0x200B)))
    })

})
