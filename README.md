# w-html2docx
A tool for docx2pdf.

![language](https://img.shields.io/badge/language-JavaScript-orange.svg) 
[![npm version](http://img.shields.io/npm/v/w-html2docx.svg?style=flat)](https://npmjs.org/package/w-html2docx) 
[![license](https://img.shields.io/npm/l/w-html2docx.svg?style=flat)](https://npmjs.org/package/w-html2docx) 
[![npm download](https://img.shields.io/npm/dt/w-html2docx.svg)](https://npmjs.org/package/w-html2docx) 
[![npm download](https://img.shields.io/npm/dm/w-html2docx.svg)](https://npmjs.org/package/w-html2docx) 
[![jsdelivr download](https://img.shields.io/jsdelivr/npm/hm/w-html2docx.svg)](https://www.jsdelivr.com/package/npm/w-html2docx)

## Documentation
To view documentation or get support, visit [docs](https://yuda-lyu.github.io/w-html2docx/global.html).

## Core
> `w-html2docx` is based on the `win32com` in `python`, and only run in `Windows`.

## Installation

### Using npm(ES6 module):
```alias
npm i w-html2docx
```

#### Example:
> **Link:** [[dev source code](https://github.com/yuda-lyu/w-html2docx/blob/master/g.mjs)]
```alias
import w from 'wsemi'
import WHtml2docx from './src/WHtml2docx.mjs'
//import WHtml2docx from 'w-html2docx/src/WHtml2docx.mjs'
//import WHtml2docx from 'w-html2docx'

async function test() {

    let fpIn = `./test/ztmp.html`
    let fpOut = `./test/ztmp.docx`
    let opt = {
        imgRatioWidthMax: 0.5,
    }

    let r = await WHtml2docx(fpIn, fpOut, opt)
    console.log(r)
    // => ok

    w.fsDeleteFile(fpOut)

}
test()
    .catch((err) => {
        console.log('catch', err)
    })
```

### 外部提供docx模板要點

轉檔時以docx模板為基底產出文件，可透過`opt.fpInTemp`指定外部模板；未指定時依序使用`./src/tmp.docx`或`./node_modules/w-html2docx/src/tmp.docx`(套件內建模板)。

外部提供模板時須注意：

1. **內容附加於模板文末**：轉檔係將html內容(含格式)插入至模板既有內容之後再另存，故模板一般提供空白文件；模板內既有內容會保留於輸出文件開頭。
2. **頁面設定由模板決定**：輸出文件之紙張大小、邊界等頁面設定沿用模板；圖片縮放上限(滿版寬高)亦依模板版心(紙張大小扣除邊界)計算。
3. **模板settings會帶入輸出文件**：模板`word/settings.xml`內之文件層級設定，轉檔後會保留於輸出docx。
4. **建議帶入停用影像壓縮旗標**：模板`word/settings.xml`須含`<w:doNotAutoCompressPictures/>`，否則Word轉檔時會將寬圖(約超過1270px)自動降採樣至220ppi並轉存JPEG，導致圖內文字模糊。套件內建模板已含此旗標；外部模板可於Word開啟模板後，勾選[檔案 > 選項 > 進階 > 影像大小和品質 > 不壓縮檔案中的影像]再存檔，或直接編輯`word/settings.xml`加入該元素。
