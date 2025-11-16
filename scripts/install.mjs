import path from 'path'
import { fileURLToPath } from 'url'
import downloadFiles from '../src/downloadFiles.mjs'


async function init() {

    //check, 被安裝條件才執行
    let __dirname = path.dirname(fileURLToPath(import.meta.url))
    if (!__dirname.includes('node_modules')) {
        return //非位於node_modules, 代表套件本身
    }

    //fdSrv
    let fdSrv = path.resolve()

    //fdBase,
    let fdBase = `${fdSrv}/src/` //npm i後觸發安裝時, 工作路徑是位於套件內
    // console.log('fdBase', fdBase)

    //downloadFiles
    await downloadFiles(fdBase)

}
init()
    .catch((err) => {
        console.log(err)
    })

//node scripts/install.mjs
