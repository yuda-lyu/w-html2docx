import fs from 'fs'
import path from 'path'


//測試用: 以repo內bin/之實檔建本機鏡像(<鏡像>/<pkg>/<sha256>/<name>, 規格O1), 並設W_BIN_MIRROR、W_BIN_CACHE、W_BIN_OFFLINE=1,
//使需要實際htmlToDocx.exe之功能測試不連外網即可取得(快取與鏡像皆在test/_tmp/<tag>/), 回傳restore函數還原環境變數並刪除該資料夾
let keys = ['W_BIN_MIRROR', 'W_BIN_CACHE', 'W_BIN_OFFLINE']


function setBinEnv(tag) {
    let fdt = path.resolve(`./test/_tmp/${tag}`)
    let manifest = JSON.parse(fs.readFileSync('bin/manifest.json', 'utf8'))

    //鏡像: 內容定址, 與快取結構相同
    for (let f of manifest.files) {
        let fd = path.join(fdt, 'mirror', manifest.pkg, f.sha256)
        fs.mkdirSync(fd, { recursive: true })
        fs.copyFileSync(path.resolve(f.path), path.join(fd, f.name))
    }

    //env
    let bk = {}
    for (let k of keys) {
        bk[k] = process.env[k]
    }
    process.env.W_BIN_MIRROR = path.join(fdt, 'mirror')
    process.env.W_BIN_CACHE = path.join(fdt, 'cache')
    process.env.W_BIN_OFFLINE = '1'

    return function restore() {
        for (let k of keys) {
            if (bk[k] === undefined) {
                delete process.env[k]
            }
            else {
                process.env[k] = bk[k]
            }
        }
        fs.rmSync(fdt, { recursive: true, force: true, maxRetries: 10, retryDelay: 300 })
        try {
            fs.rmdirSync(path.resolve('./test/_tmp')) //尚有其他測試之暫存時非空, 略過
        }
        catch (err) {}
    }
}


export { setBinEnv }
