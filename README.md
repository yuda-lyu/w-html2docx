# w-html2docx
A tool for docx2pdf.

w-html2docx w-html2docx
WHtml2docx WHtml2docx

![language](https://img.shields.io/badge/language-JavaScript-orange.svg) 
[![npm version](http://img.shields.io/npm/v/w-html2docx.svg?style=flat)](https://npmjs.org/package/w-html2docx) 
[![license](https://img.shields.io/npm/l/w-html2docx.svg?style=flat)](https://npmjs.org/package/w-html2docx) 
[![npm download](https://img.shields.io/npm/dt/w-html2docx.svg)](https://npmjs.org/package/w-html2docx) 
[![npm download](https://img.shields.io/npm/dm/w-html2docx.svg)](https://npmjs.org/package/w-html2docx) 
[![jsdelivr download](https://img.shields.io/jsdelivr/npm/hm/w-html2docx.svg)](https://www.jsdelivr.com/package/npm/w-html2docx)

## Documentation
To view documentation or get support, visit [docs](https://yuda-lyu.github.io/w-html2docx/global.html).

## Core
> `w-html2docx` is based on the `docx2pdf` in `python`, and only run in `Windows`.

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
    let opt = {}

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
