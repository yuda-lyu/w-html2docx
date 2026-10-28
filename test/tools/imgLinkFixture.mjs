import fs from 'fs'
import path from 'path'
import { pathToFileURL } from 'url'
import { unzipSync, strFromU8 } from 'fflate'
import { genPng } from './imgHeightFixture.mjs'


//連結圖測試之夾具與拆包工具: 10案圖片引用寫法(路徑含中文與空白)、html、拆docx列圖片關聯


//JPG_B64, 240×120 px之JPEG(漸層, 品質80, 1737位元組; 以sharp一次性產生後內嵌, 使測試不依賴影像套件)
let JPG_B64 = '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'

//SVG, 240×120之向量圖
let SVG = '<svg xmlns="http://www.w3.org/2000/svg" width="240" height="120" viewBox="0 0 240 120"><rect width="240" height="120" fill="#fde68a"/><circle cx="60" cy="60" r="40" fill="#2563eb"/><rect x="120" y="30" width="90" height="60" fill="#dc2626"/></svg>'


//LINK_CASES, 10案圖片引用寫法(同審計md2docx-run01之E2, 其中S0改為JPEG之data URI)
//  kind: embed為Word可解析之寫法(轉檔後須內嵌), unresolved為Word無法解析之寫法(轉檔須失敗並列出路徑), data為data URI(照舊內嵌)
let LINK_CASES = [
    { id: 'A1', desc: '絕對路徑(反斜線)', file: '圖 A1.png', kind: 'embed' },
    { id: 'A2', desc: '絕對路徑(正斜線)', file: '圖 A2.png', kind: 'unresolved' },
    { id: 'R1', desc: '相對路徑(原字)', file: '圖 R1.png', kind: 'embed' },
    { id: 'R2', desc: '相對路徑(百分比編碼)', file: '圖 R2.png', kind: 'unresolved' },
    { id: 'F1', desc: 'file:///網址(百分比編碼, pathToFileURL)', file: '圖 F1.png', kind: 'unresolved' },
    { id: 'F2', desc: 'file:///網址(原字)', file: '圖 F2.png', kind: 'embed' },
    { id: 'J1', desc: '絕對路徑 JPG', file: '圖 J1.jpg', kind: 'embed' },
    { id: 'S1', desc: '絕對路徑 SVG', file: '圖 S1.svg', kind: 'embed' },
    { id: 'B', desc: 'data URI(PNG)', file: '圖 B.png', kind: 'data', mime: 'image/png' },
    { id: 'D2', desc: 'data URI(JPEG)', file: '圖 D2.jpg', kind: 'data', mime: 'image/jpeg' },
]


//genLinkCases, 於fd下建圖片夾「圖 資料夾」與網頁夾「頁 面」(含中文與空白), 回傳LINK_CASES各案(另有fp、bytes、src)與網頁夾
//  fp為圖檔絕對路徑(反斜線); bytes為圖檔內容, 內嵌後word/media內須有位元組相同之檔(data URI者為其解碼內容); src為html內img之src
function genLinkCases(fd) {
    let fdImg = path.join(fd, '圖 資料夾')
    let fdHtml = path.join(fd, '頁 面')
    fs.mkdirSync(fdImg, { recursive: true })
    fs.mkdirSync(fdHtml, { recursive: true })

    let png = (rgb) => genPng(320, 180, rgb)
    let jpg = Buffer.from(JPG_B64, 'base64')
    let svg = Buffer.from(SVG, 'utf8')
    let bytesOf = {
        A1: png([230, 60, 60]),
        A2: png([60, 160, 60]),
        R1: png([60, 60, 230]),
        R2: png([230, 160, 30]),
        F1: png([160, 60, 200]),
        F2: png([30, 190, 190]),
        J1: jpg,
        S1: svg,
        B: png([120, 120, 120]),
        D2: jpg,
    }

    let cs = LINK_CASES.map((v) => ({ ...v, bytes: bytesOf[v.id] }))

    for (let c of cs) {
        c.fp = path.join(fdImg, c.file)
        fs.writeFileSync(c.fp, c.bytes)
        let rel = path.relative(fdHtml, c.fp).replace(/\\/g, '/')
        let src = ''
        if (c.id === 'A1' || c.id === 'J1' || c.id === 'S1') {
            src = c.fp
        }
        else if (c.id === 'A2') {
            src = c.fp.replace(/\\/g, '/')
        }
        else if (c.id === 'R1') {
            src = rel
        }
        else if (c.id === 'R2') {
            src = rel.split('/').map((s) => (s === '..' ? s : encodeURIComponent(s))).join('/')
        }
        else if (c.id === 'F1') {
            src = pathToFileURL(c.fp).href
        }
        else if (c.id === 'F2') {
            src = `file:///${c.fp.replace(/\\/g, '/')}`
        }
        else {
            src = `data:${c.mime};base64,${c.bytes.toString('base64')}`
        }
        c.src = src
    }

    return { cases: cs, fdHtml }
}


//genHtml, 以body內容組出html存於fp
function genHtml(body, fp) {
    let html = `<!DOCTYPE html>\n<html>\n<head>\n<meta charset="utf-8" />\n<title>imgLink</title>\n</head>\n<body>\n${body}\n</body>\n</html>\n`
    fs.writeFileSync(fp, html, 'utf8')
    return fp
}


//genLinkHtml, 以案組出html(每案一段說明與一張圖), 存於fp
function genLinkHtml(cases, fp) {
    return genHtml(cases.map((c) => `<p>CASE ${c.id}: ${c.desc}</p>\n<p><img src="${c.src}" alt="${c.id}"></p>`).join('\n'), fp)
}


//attr, 取標籤字串內屬性值
function attr(tag, name) {
    let m = tag.match(new RegExp(`\\s${name}="([^"]*)"`))
    return m ? m[1] : ''
}


//getDocxPictures, 拆docx回傳圖片相關資訊:
//  media: word/media下各檔{ name, bytes }
//  blips: word/document.xml內各圖片參照(a:blip與asvg:svgBlip){ embed, link }(屬性值為關聯id, 無則為空字串)
//  imageRels: word/_rels/document.xml.rels內類型為image之關聯{ id, target, external }
//  extents: 各行內圖片之顯示尺寸(wp:extent), 單位點
function getDocxPictures(fp) {
    let files = unzipSync(new Uint8Array(fs.readFileSync(fp)))
    let doc = strFromU8(files['word/document.xml'])
    let rels = strFromU8(files['word/_rels/document.xml.rels'])

    let media = Object.keys(files).filter((k) => k.startsWith('word/media/')).map((k) => ({ name: k.slice('word/'.length), bytes: Buffer.from(files[k]) }))

    let blips = [...doc.matchAll(/<(?:a:blip|asvg:svgBlip)\b[^>]*>/g)].map((m) => ({ embed: attr(m[0], 'r:embed'), link: attr(m[0], 'r:link') }))

    let imageRels = [...rels.matchAll(/<Relationship\b[^>]*>/g)].map((m) => m[0]).filter((t) => /\/relationships\/image"/.test(t)).map((t) => ({
        id: attr(t, 'Id'),
        target: attr(t, 'Target'),
        external: attr(t, 'TargetMode') === 'External',
    }))

    let extents = [...doc.matchAll(/<wp:extent cx="(\d+)" cy="(\d+)"/g)].map((m) => ({ cx: Number(m[1]) / 12700, cy: Number(m[2]) / 12700 }))

    return { media, blips, imageRels, extents }
}


export {
    JPG_B64,
    SVG,
    LINK_CASES,
    genLinkCases,
    genHtml,
    genLinkHtml,
    getDocxPictures
}
