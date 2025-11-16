import fs from 'fs'
import w from 'wsemi'
import assert from 'assert'
import WHtml2docx from '../src/WHtml2docx.mjs'


function isWindows() {
    return process.platform === 'win32'
}


describe('WHtml2docx', function() {

    //check
    if (!isWindows()) {
        return
    }

    let fpOutTrue = `./test/ztmpTrue.docx`

    let fpIn = `./test/ztmp.html`
    let fpOut = `./test/ztmp.docx`
    let opt = {
        imgRatioWidthMax: 0.5,
    }

    it('convert', async function() {
        await WHtml2docx(fpIn, fpOut, opt)
        let r = (fs.statSync(fpOut)).size
        let rr = (fs.statSync(fpOutTrue)).size
        //轉出docx檔案每次不同, 改用門檻比對
        // console.log('r', r, 'rr', rr)
        let b = r > 62000 && rr > 62000
        w.fsDeleteFile(fpOut)
        assert.strict.deepEqual(true, b)
    })

})
