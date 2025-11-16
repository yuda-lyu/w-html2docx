import path from 'path'
import process from 'process'
import get from 'lodash-es/get.js'
import isestr from 'wsemi/src/isestr.mjs'
import isearr from 'wsemi/src/isearr.mjs'
import cstr from 'wsemi/src/cstr.mjs'
import str2b64 from 'wsemi/src/str2b64.mjs'
import execProcess from 'wsemi/src/execProcess.mjs'
import fsIsFile from 'wsemi/src/fsIsFile.mjs'


let fdSrv = path.resolve()


function isWindows() {
    return process.platform === 'win32'
}


/**
 * Html檔轉Docx檔
 *
 * @param {String} fpInHtml 輸入來源Html檔位置字串
 * @param {String} fpOutDocx 輸入轉出Docx檔位置字串
 * @param {Object} [opt={}] 輸入設定物件，預設{}
 * @param {Array|String} [opt.fontFamilies=['標楷體', 'Times New Roman']] 輸入Docx更改字型陣列或字串，預設['標楷體', 'Times New Roman']
 * @returns {Promise} 回傳Promise，resolve回傳成功訊息，reject回傳錯誤訊息
 * @example
 *
 * import w from 'wsemi'
 * import WHtml2docx from './src/WHtml2docx.mjs'
 * //import WHtml2docx from 'w-html2docx/src/WHtml2docx.mjs'
 * //import WHtml2docx from 'w-html2docx'
 *
 * async function test() {
 *
 *     let fpIn = `./test/ztmp.html`
 *     let fpOut = `./test/ztmp.docx`
 *     let opt = {}
 *
 *     let r = await WHtml2docx(fpIn, fpOut, opt)
 *     console.log(r)
 *     // => ok
 *
 *     w.fsDeleteFile(fpOut)
 *
 * }
 * test()
 *     .catch((err) => {
 *         console.log('catch', err)
 *     })
 *
 */
async function WHtml2docx(fpInHtml, fpOutDocx, opt = {}) {
    let errTemp = null

    //isWindows
    if (!isWindows()) {
        return Promise.reject('operating system is not windows')
    }

    //check
    if (!fsIsFile(fpInHtml)) {
        return Promise.reject(`fpInHtml[${fpInHtml}] does not exist`)
    }

    //轉絕對路徑
    fpInHtml = path.resolve(fpInHtml)
    fpOutDocx = path.resolve(fpOutDocx)

    //fontFamilies
    let fontFamilies = get(opt, 'fontFamilies')
    if (isestr(fontFamilies)) {
        fontFamilies = [fontFamilies]
    }
    if (!isearr(fontFamilies)) {
        fontFamilies = ['標楷體', 'Times New Roman']
    }

    //fnTmp
    let fnTmp = `tmp.docx`

    //fpTmp
    let fpTmp = ``
    if (true) {
        let fpTmpSrc = `${fdSrv}/src/${fnTmp}`
        let fpTmpNM = `${fdSrv}/node_modules/w-html2docx/src/${fnTmp}`
        if (fsIsFile(fpTmpSrc)) {
            fpTmp = fpTmpSrc
        }
        else if (fsIsFile(fpTmpNM)) {
            fpTmp = fpTmpNM
        }
        else {
            return Promise.reject(`can not find folder for ${fnTmp}`)
        }

        //轉絕對路徑
        fpTmp = path.resolve(fpTmp)

    }

    //fnExe
    let fnExe = `htmlToDocx.exe`

    //fdExe
    let fdExe = ''
    if (true) {
        let fdExeSrc = `${fdSrv}/src/`
        let fdExeNM = `${fdSrv}/node_modules/w-html2docx/src/`
        if (fsIsFile(`${fdExeSrc}${fnExe}`)) {
            fdExe = fdExeSrc
        }
        else if (fsIsFile(`${fdExeNM}${fnExe}`)) {
            fdExe = fdExeNM
        }
        else {
            return Promise.reject('can not find folder for html2docx')
        }
    }
    // console.log('fdExe', fdExe)

    //prog
    let prog = `${fdExe}${fnExe}`
    // console.log('prog', prog)

    //inp
    let inp = {
        fpInSrc: fpInHtml,
        fpInTemp: fpTmp,
        fpOut: fpOutDocx,
        fontFamilies: ['標楷體', 'Times New Roman'],
    }
    // console.log('inp', inp)

    //input to b64
    let cInput = JSON.stringify(inp)
    let b64Input = str2b64(cInput)
    // console.log('b64Input', b64Input)

    //execProcess
    await execProcess(prog, b64Input)
        .catch((err) => {
            console.log('execProcess catch', err)
            errTemp = err.toString()
        })

    //check
    if (errTemp) {
        return Promise.reject(errTemp)
    }

    // //check
    // if (!isestr(output)) {
    //     return Promise.reject(`output[${cstr(output)}] is not an effective string`)
    // }

    return 'ok'
}


export default WHtml2docx
