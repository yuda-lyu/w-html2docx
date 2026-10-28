import fs from 'fs'
import path from 'path'
import zlib from 'zlib'
import { execFileSync } from 'child_process'
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate'


//imgHeightReserveLines測試之夾具與量測工具: 長圖PNG、行格線模板、圖區塊html、讀圖高、以Word量圖與圖名之頁碼


//crc32, PNG區塊校驗碼
//why: 不依賴zlib.crc32(Node 22.2起才有)
let crcTable = null
function crc32(buf) {
    if (crcTable === null) {
        crcTable = []
        for (let n = 0; n < 256; n++) {
            let c = n
            for (let k = 0; k < 8; k++) {
                c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1)
            }
            crcTable.push(c >>> 0)
        }
    }
    let c = 0xffffffff
    for (let i = 0; i < buf.length; i++) {
        c = crcTable[(c ^ buf[i]) & 0xff] ^ (c >>> 8)
    }
    return (c ^ 0xffffffff) >>> 0
}


//genPng, 產生單色RGB之PNG(w×h px, rgb為[r,g,b]), 色不同則位元組不同, 供需辨識是哪一張圖之測試使用
function genPng(w, h, rgb) {
    let row = Buffer.alloc(1 + w * 3)
    row[0] = 0 //filter none
    for (let x = 0; x < w; x++) {
        row[1 + x * 3] = rgb[0]
        row[2 + x * 3] = rgb[1]
        row[3 + x * 3] = rgb[2]
    }
    let raw = Buffer.concat(Array.from({ length: h }, () => row))
    let chunk = (type, data) => {
        let len = Buffer.alloc(4)
        len.writeUInt32BE(data.length)
        let td = Buffer.concat([Buffer.from(type, 'ascii'), data])
        let crc = Buffer.alloc(4)
        crc.writeUInt32BE(crc32(td))
        return Buffer.concat([len, td, crc])
    }
    let ihdr = Buffer.alloc(13)
    ihdr.writeUInt32BE(w, 0)
    ihdr.writeUInt32BE(h, 4)
    ihdr[8] = 8 //bit depth
    ihdr[9] = 2 //RGB
    let sig = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
    return Buffer.concat([sig, chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))])
}


//genTallPng, 產生單色RGB之PNG, 預設1200×2800px, 模擬整頁高之長圖(網站設計稿、長流程圖)
function genTallPng(w = 1200, h = 2800) {
    return genPng(w, h, [220, 220, 220])
}


//genGridTemplate, 由內建模板產生行格線模板: 文件格線改為只指定行(行距18點, A4上下邊界72點時每頁38格), 並移除內文樣式之段前段後距離,
//使單行之段落恰佔1格(與w-md2docx之temp_tpc.docx同構); 另移除頁首頁尾之參照:
//內建模板之頁尾(2段、各帶段前段後6點)高於下邊界與頁尾距離之差, Word會縮小本文區, 實測每頁僅37格, 與版心高推算之38格不符
function genGridTemplate(fpSrc, fpOut) {
    let files = unzipSync(new Uint8Array(fs.readFileSync(fpSrc)))
    let doc = strFromU8(files['word/document.xml'])
    let st = strFromU8(files['word/styles.xml'])
    let doc2 = doc.replace(/<w:docGrid\b[^>]*\/>/, '<w:docGrid w:type="lines" w:linePitch="360"/>').replace(/<w:(header|footer)Reference\b[^>]*\/>/g, '')
    let st2 = st.replace(/(<w:style w:type="paragraph" w:default="1"[\s\S]*?<w:pPr>)([\s\S]*?)(<\/w:pPr>)/, (m, a, b, c) => a + b.replace(/<w:spacing\b[^>]*\/>/, '') + c)
    if (doc2 === doc || st2 === st) {
        throw new Error('can not derive the line grid template from the built-in template')
    }
    fs.writeFileSync(fpOut, zipSync({ ...files, 'word/document.xml': strToU8(doc2), 'word/styles.xml': strToU8(st2) }))
}


//genFigureHtml, 前言段落後接n個圖區塊(圖片+圖名), 圖區塊間以w-md2html之換行標記(內含零寬空格之空div)間隔, 圖名於版心寬415.3點、12pt時折成2行
//前言與結尾用div而不用p: p轉入後為網頁內文樣式, 帶段前段後各5點, 使其後各行不再對齊格線而多佔不足1格, 頁面格數即無法精確推算
//  如此第1頁為前言1格+圖+圖名2格, 其後各頁為間隔段1格(被擠至頁首)+圖+圖名2格
function genFigureHtml(b64, n = 4) {
    let zwsp = '&#8203;'
    let blocks = []
    for (let i = 1; i <= n; i++) {
        blocks.push(`<div style="padding:2rem 0; text-align:center;"><img src="data:image/png;base64,${b64}" style="width:100%; max-width:700px;"><div style="padding-top:0.5rem;">圖${i} 用過核子燃料最終處置資訊展示網站之首頁設計稿，含導覽列、主視覺、最新消息、計畫沿革、各階段成果展示與頁尾聯絡資訊之完整版面配置示意</div></div>`)
        if (i < n) {
            blocks.push(`<div style="text-align:center;" data-for="br">${zwsp}</div>`)
        }
    }
    return `<html><head><meta charset="utf-8"></head><body><div>前言段落。</div>${blocks.join('')}<div>結尾段落。</div></body></html>`
}


//getImgHeights, 讀docx內各行內圖片之高度(點), 由word/document.xml之wp:extent cy(EMU)換算
function getImgHeights(fp) {
    let xml = strFromU8(unzipSync(new Uint8Array(fs.readFileSync(fp)))['word/document.xml'])
    return [...xml.matchAll(/<wp:extent cx="(\d+)" cy="(\d+)"/g)].map((m) => Number(m[2]) / 12700)
}


//getFigurePages, 以Word開啟docx(唯讀、不顯示), 回傳各圖片段之頁碼與其後第1個非空段落(圖名)首行、末行之頁碼
function getFigurePages(fp) {
    let fpVbs = path.resolve('./test/tools/figurePages.vbs')
    let out = execFileSync('cscript', ['//nologo', fpVbs, path.resolve(fp)], { encoding: 'utf8', windowsHide: true, timeout: 120000 })
    return out.trim().split(/\r?\n/).filter((v) => v !== '').map((v) => {
        let [img, capStart, capEnd] = v.split(',').map(Number)
        return { img, capStart, capEnd }
    })
}


export {
    genPng,
    genTallPng,
    genGridTemplate,
    genFigureHtml,
    getImgHeights,
    getFigurePages
}
