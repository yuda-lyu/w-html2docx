import path from 'path'
import { fileURLToPath } from 'url'
import autoDownloadFiles from '../src/autoDownloadFiles.mjs'


async function init() {

    //check
    let __dirname = path.dirname(fileURLToPath(import.meta.url))
    if (!__dirname.includes('node_modules')) {
        return //非位於node_modules, 代表套件本身
    }

    //autoDownloadFiles, postinstall時cwd=套件自身在node_modules內的目錄, 其依序查找<cwd>/src/與<cwd>/node_modules/w-html2docx/src/, 皆無則下載至<cwd>/src/,
    //與執行期缺檔時之自動下載為同一套邏輯, 不另行組路徑; 已存在(重新安裝)則不重複下載
    let { fpExe } = await autoDownloadFiles()
    console.log(`htmlToDocx.exe is ready at [${fpExe}]`)

}
init()
    .catch((err) => {
        console.log(err)
    })

//node scripts/install.mjs
