# 數學方程式編輯器（Word / Excel 增益集）

給出考卷、講義用的數學方程式編輯器。插入的是 **Word 內建方程式**，不是圖片，也不是 MathType 物件。

| 常見問題 | 這個工具 |
|---|---|
| Word 方程式寫在句子裡，**分數會自動縮小** | 行內分數維持正常大小（可關閉） |
| MathType 方程式在**沒裝 MathType 的電腦**只能當圖片、不能改 | 插入 Word 原生方程式，任何電腦打開都能直接編輯 |
| MathType 物件排版時**往上飄**、跟文字對不齊 | 原生方程式跟文字同一條基準線 |
| 要分享給同事 | 開源、免安裝程式，同事匯入一個設定檔就能用 |

支援 **Word / Excel 2016、2019、2021、Microsoft 365（Windows、Mac）與網頁版**。

## 功能

- **所見即所得輸入**：打 `/` 變分數、`^` 上標、`_` 下標，也有按鈕面板（根號、向量、線段、∠、△、°、Σ、∫、lim、log、組合數…）。
- **分數不縮小**：Word 會把行內分數縮成約 67%，這裡把分子、分母字級放大抵銷（實測 1.49 倍）。仍然是可編輯的 Word 方程式。
- **讀回修改**：選取文件裡的方程式按「讀取選取的方程式」，就能在編輯器修改後取代——**用 Word 內建方程式（Alt+=）打的也可以**。
- **一鍵整理舊文件**：選取範圍（或 Ctrl+A 全選）按「放大選取範圍的分數」，既有的行內分數全部變正常大小；也能還原。
- **矩陣、聯立方程式、分段函數**：一格一個小編輯框，「＋列／－列／＋欄／－欄」增刪，Enter 加一列、空白列按 Backspace 刪掉，「刪除整個」一次清掉。
- **其他進階結構**：弧（⌒AB）、方框、上方加字等，用 LaTeX 輸入。
- **Excel**：複製後在 Excel 的方程式裡貼上、按「專業」，變成 Excel 原生方程式。
- **最近使用**：插入過的式子點一下就能再用。

## 安裝（給老師）

增益集的程式放在網路上（GitHub Pages），電腦只需要匯入 `manifest.xml` 這個設定檔。**使用時需要連網。**

> 下面的 `manifest.xml` 是指**已經改成你們網址的版本**（見下方「架設自己的版本」）。

### 方法一：Word / Excel 網頁版（最簡單）

1. 開啟 Word 網頁版的任一文件。
2. **常用 → 增益集 → 進階（或「更多增益集」）→ 我的增益集 → 上傳我的增益集**。
3. 選 `manifest.xml` 上傳。

### 方法二：Windows 桌面版（2016 / 2019 / 365）— 共用資料夾

Office 2016 以後的 Windows 桌面版，要從「受信任的增益集目錄」安裝。一間辦公室只要一台電腦提供共用資料夾即可。

**在一台電腦上（只做一次）：**

1. 建立資料夾，例如 `C:\OfficeAddins`，把 `manifest.xml` 放進去。
2. 對資料夾按右鍵 → **內容 → 共用 → 進階共用** → 勾選「共用此資料夾」。
3. 記下網路路徑，例如 `\\出題電腦\OfficeAddins`。

**每位老師的電腦：**

1. 開啟 Word → **檔案 → 選項 → 信任中心 → 信任中心設定 → 受信任的增益集目錄**。
2. 「目錄 URL」填入 `\\出題電腦\OfficeAddins` → **新增目錄** → 勾選 **在功能表中顯示** → 確定。
3. **關閉並重新開啟 Word**。
4. **插入 → 我的增益集**（或「增益集 → 更多增益集」）→ 上方切到 **共用資料夾** → 選「數學方程式編輯器」→ 新增。
5. 之後在 **常用** 索引標籤最右邊會有「數學方程式」按鈕。Excel 設定一次後也能用。

### 方法三：學校有 Microsoft 365 系統管理員

請管理員在 **Microsoft 365 系統管理中心 → 設定 → 整合式應用程式 → 上傳自訂應用程式** 上傳 `manifest.xml`，指派給老師即可，老師不用自己設定。

## 使用方式

### Word

1. 把游標放在要插入的位置。
2. 在編輯器打式子（或點按鈕），按 **插入**（或 Enter）。
3. 要修改：**拖曳選取**文件中的方程式 → **讀取選取的方程式** → 改好按 **插入**（會取代選取的方程式）。

選項：

- **分數不縮小（行內）**：預設開啟。
- **獨立一行置中**：插入成單獨一行的方程式（顯示模式）。

### Excel

1. 在編輯器打好式子，按 **複製到 Excel**。
2. 在 Excel 按 `Alt` + `=`（插入 → 方程式）。
3. `Ctrl` + `V` 貼上。
4. 在「方程式」索引標籤按 **專業**（或對方程式按右鍵 → 專業）。

### 進階：LaTeX

展開「LaTeX 原始碼」可以直接打 LaTeX。分段函數、矩陣等視覺編輯器不支援的結構，在「進階」分頁點一下就會放進 LaTeX 範本。

## 已知限制

- **行內的 Σ、∏**：Word 規定行內時上下限放在右側；要放在符號上下方請勾「獨立一行置中」。
- **放大分數是調整字級**：之後若把整段改成別的字級，請再按一次「放大選取範圍的分數」。
- **視覺編輯器**（MathQuill）不支援弧、方框、上方加字等，這些請用 LaTeX；轉換器本身都支援。一個式子只能放一個矩陣／大括號（要放兩個請用 LaTeX）。
- 不支援 Office 2013 以前的版本。

## 架設自己的版本（給管理者）

1. 在 GitHub **Fork** 這個專案（或建立新 repository 上傳全部檔案）。
2. 到 repository 的 **Settings → Pages**，Source 選 `Deploy from a branch`、Branch 選 `main` / `(root)`，儲存。
   幾分鐘後網址會是 `https://你的帳號.github.io/office-math-editor/`。
3. 把 manifest 裡的網址改成你的：

   ```bash
   node tools/set-base-url.js https://你的帳號.github.io/office-math-editor
   ```

   （沒有 Node.js 的話，用記事本把 `manifest.xml` 裡所有 `YOUR-GITHUB-NAME` 換成你的帳號也可以。）
4. Commit、push，把新的 `manifest.xml` 發給老師。

## 開發與測試

```bash
npm install
npm test                                              # 101 項：ES5 語法、來回轉換、Word 原生方程式、分數放大、段落插入
node tests/make-docx.js out.json && python tests/render.py out.json 輸出資料夾   # 用 Word 實際排版檢查（Windows + Word）
powershell -File tests/word-flow.ps1 輸出資料夾          # 模擬增益集在 Word 的插入與整理流程
```

程式全部是 **ES5**：Office 2016/2019（非訂閱版）的增益集跑在 IE11 核心，不能用箭頭函式、`let`、Promise 語法糖等；`npm test` 會檢查。

| 檔案 | 用途 |
|---|---|
| `src/latex2omml.js` | LaTeX → Word 方程式（OMML）、UnicodeMath 線性格式 |
| `src/omml2latex.js` | Word 方程式 → LaTeX（讀回修改） |
| `src/insert-inline.js` | 在句子中間插入方程式而不斷行 |
| `src/fixfrac.js` | 放大既有文件的行內分數 |
| `src/taskpane.*` | 工作窗格介面 |
| `manifest.xml` | 增益集設定檔 |

## 授權

MIT License。使用的第三方程式（執行時從 CDN 載入，未包含在本專案）：

- [MathQuill](https://github.com/mathquill/mathquill)（MPL-2.0）
- [jQuery](https://jquery.com/)（MIT）
- Office JavaScript API（Microsoft）
