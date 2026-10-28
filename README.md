# w-html2docx
A tool for html2docx.

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
5. **行格線模板之圖名預留**：模板設有文件格線(如[版面配置 > 版面設定 > 文件格線 > 只指定行的格線])時，Word會把每段之高度進位為行距之整數倍，整頁高之圖片進位後常擠掉其下之圖名。此時可給`opt.imgHeightReserveLines`(預留格數)，圖片最大高度改為`(每頁格數－預留格數)×行距－0.5點`，行距取模板文件格線之行距，每頁格數為版心高÷行距之整數部分。例如行距18點、A4上下邊界各72點之模板每頁38格，圖區塊間有1行間隔段、圖名至多2行時至少給3，建議給4多留1格餘裕；14pt之標楷體每行佔2格，須依實際所佔格數加計。模板無文件格線時此設定不生效，圖片最大高度維持版心高×0.937043054427296×`opt.imgRatioHeightMax`。

## 圖片

html內的圖片皆內嵌於輸出docx，轉出的docx不依賴原圖(移至他處或刪除原圖後仍可顯示)：

- `data:`網址(base64)之圖片：照舊內嵌。
- 本機圖片(`<img src>`指向本機檔案)：Word匯入html時本機圖片為連結圖，轉檔時逐一斷開連結並將圖片存入docx。png、jpg維持原格式與位元組，svg以向量(svg)加點陣預覽(png)內嵌。
- 網址圖片(`http://`、`https://`)：由Word於轉檔時取回後內嵌，故轉檔當下須連得到。
- 浮動圖片(如`<img align="right">`，Word匯入後為浮動圖片)：同上。
- 解析不到來源之圖片：不產生輸出檔，`WHtml2docx()`以字串reject，每行列出一個無法解析之來源，例如：

```alias
code[1]:
unresolved linked picture: C://opensrc/專案/圖 資料夾/圖 A2.png
error: LinkedPictureError: 1 linked picture(s) can not be embedded since Word can not resolve their source(see stderr)
```

本機路徑之寫法依Word決定(Word 16實測，路徑可含中文與空白)，須用下表**可解析**之寫法：

| 寫法 | 範例 | 結果 |
|---|---|---|
| 反斜線絕對路徑 | `C:\專案\圖 資料夾\a.png` | 可解析 |
| 相對路徑(原字，相對於html檔所在資料夾) | `../圖 資料夾/a.png` | 可解析 |
| `file:///`網址(原字，不做百分比編碼) | `file:///C:/專案/圖 資料夾/a.png` | 可解析 |
| 正斜線絕對路徑 | `C:/專案/圖 資料夾/a.png` | **不可解析**(Word改寫為`C://專案/...`) |
| 百分比編碼之相對路徑 | `../%E5%9C%96%20.../a.png` | **不可解析** |
| 百分比編碼之`file:///`網址(如`pathToFileURL()`之結果) | `file:///C:/%E5%B0%88.../a.png` | **不可解析** |

不可解析之寫法在Word內成為24×24之破圖連結；亦可改用`data:`網址避開路徑問題。

## Breaking changes

本次調整(`1.0.39`之後的版本)：

1. **本機圖片與網址圖片改為內嵌**：改前Word匯入html之本機圖片存為外部連結(`TargetMode="External"`，`<a:blip>`只有`r:link`)，docx須與原圖同處才顯示，移走即空框；改後存入docx(`word/media/`，`r:embed`)，docx不再依賴原圖，檔案隨圖片增大。網址圖片(`http://`、`https://`)改前亦存為連結，改後由Word取回內嵌。
2. **解析不到之圖片改為轉檔失敗**：改前此類圖片悄悄存為24×24之破圖連結且轉檔成功；改後`WHtml2docx()`reject並列出來源，不產生輸出檔。html若含正斜線絕對路徑、百分比編碼之路徑或網址，須改用上表可解析之寫法(或`data:`網址)。
3. **轉檔器`htmlToDocx.exe`更新**：以上行為由轉檔器實作；建置輸出改至`bin/htmlToDocx.exe`(`srcPython/build.bat`)，`src/htmlToDocx.exe`為舊版相容保留檔，不再更新。

## 外部執行檔

本套件轉檔時需要下列外部執行檔，不含於npm包：

| 檔名 | 用途 | 來源 |
|---|---|---|
| htmlToDocx.exe | 調用本機Microsoft Word(win32com)將Html轉為Docx | 本套件GitHub repo之`bin/`(釘選提交) |

- 首次轉檔時自動下載(或預先呼叫本套件匯出之`ensureBins()`)，以`bin/manifest.json`所載之SHA-256與大小校驗後存於`<套件根>/bin-cache/`；之後自快取取得(亦校驗)。安裝時不下載。
- 環境變數：
  - `W_BIN_CACHE=<資料夾>`：快取改放`<資料夾>/w-html2docx/`(套件所在之`node_modules`唯讀時使用，多套件可共用同一資料夾)。
  - `W_BIN_MIRROR=<資料夾或http(s)網址>`：內網鏡像，先找鏡像再連GitHub；結構為`<鏡像>/w-html2docx/<sha256>/htmlToDocx.exe`。鏡像之檔同樣以清單校驗，不符即不用。
  - `W_BIN_OFFLINE=1`：不連網，只用快取與鏡像；缺檔時錯誤(`E_OFFLINE_MISSING`)會列出缺哪些檔、sha256、大小與應放之完整路徑。
- 離線部署：
  1. 在可連網之電腦設`W_BIN_CACHE=<空資料夾>`，呼叫本套件之`ensureBins()`一次(或轉檔一次)。
  2. 把該資料夾整個帶到內網電腦，設`W_BIN_MIRROR=<該資料夾>`(全離線時另設`W_BIN_OFFLINE=1`)。
- `WHtml2docx()`取不到執行檔時reject字串(含錯誤代碼，如`[E_OFFLINE_MISSING]`、`[E_HASH_MISMATCH]`)；直接呼叫`ensureBins()`(`WHtml2docx.ensureBins`或`w-html2docx/src/ensureBins.mjs`)則reject帶`code`之Error。

```alias
import WHtml2docx from 'w-html2docx'

//部署前預先取得, 回傳{ 'htmlToDocx.exe': 絕對路徑 }
let bins = await WHtml2docx.ensureBins()
console.log(bins['htmlToDocx.exe'])
```
