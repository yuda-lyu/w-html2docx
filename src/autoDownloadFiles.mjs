import ensureBins from './ensureBins.mjs'


//fnExe, htmlToDocx執行檔名稱(bin/manifest.json之name)
let fnExe = 'htmlToDocx.exe'


/**
 * 取得htmlToDocx.exe的絕對路徑，保留作為舊版對外匯入點之相容層，實際改由ensureBins取得
 *
 * 舊版於當前工作路徑的src/與node_modules/w-html2docx/src/查找，缺檔時自主幹下載且不校驗；現改經ensureBins：首次使用時自GitHub(釘選提交)或環境變數W_BIN_MIRROR之鏡像取得，以bin/manifest.json之SHA-256與大小校驗後存於快取(預設為套件根之bin-cache/，可由W_BIN_CACHE改位置)，之後自快取取得(亦校驗)，故回傳之路徑位於快取而非src/。不在安裝時下載
 *
 * 新程式碼請直接調用ensureBins(`w-html2docx/src/ensureBins.mjs`或`WHtml2docx.ensureBins`)
 *
 * 因htmlToDocx.exe只能用於Windows作業系統，故調用前須自行檢核作業系統
 *
 * @deprecated 改用ensureBins
 * @returns {Promise} 回傳Promise，resolve回傳執行檔路徑物件，內含fpExe為htmlToDocx.exe的絕對路徑字串，reject回傳帶code之Error(如E_OFFLINE_MISSING、E_HASH_MISMATCH)
 * @example
 * import autoDownloadFiles from 'w-html2docx/src/autoDownloadFiles.mjs'
 *
 * async function test() {
 *
 *     //autoDownloadFiles, 經ensureBins取得並校驗htmlToDocx.exe, 取不到則reject
 *     let { fpExe } = await autoDownloadFiles()
 *
 *     console.log('fpExe', fpExe)
 *     // fpExe D:\xxx\node_modules\w-html2docx\bin-cache\<sha256>\htmlToDocx.exe
 * }
 * test()
 *     .catch((err) => {
 *         console.log('catch', err)
 *     })
 *
 */
async function autoDownloadFiles() {
    let bins = await ensureBins({ names: fnExe })
    return {
        fpExe: bins[fnExe],
    }
}


export default autoDownloadFiles
