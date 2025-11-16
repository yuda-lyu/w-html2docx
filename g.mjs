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


//node g.mjs
