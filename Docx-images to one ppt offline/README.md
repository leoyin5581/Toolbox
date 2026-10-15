# DOCX 圖片提取 → PPT 生成器 v3

> 從多個 Word (.docx) 文件中批量提取圖片，自動排成網格總覽頁，一鍵生成 PowerPoint 簡報。

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![No Server](https://img.shields.io/badge/server-none-green.svg)
![Offline](https://img.shields.io/badge/offline-ready-green.svg)

## 功能特色

- **多檔案批量處理** — 一次拖入多個 .docx，每個檔案獨立處理
- **自動網格排版** — 同一個 Word 的所有圖片自動排成整齊的網格，塞在同一張 slide
- **批量排除規則** — 可設定每個檔案自動排除前 N 張圖片（如頁首頁尾 logo）
- **完全離線** — 單檔 HTML，所有函式庫已內嵌，不需網路連線
- **純瀏覽器運行** — 不需要 Python、Node.js 或任何後端環境
- **隱私安全** — 所有處理都在瀏覽器本地端完成，檔案不會上傳到任何伺服器
- **互動式預覽** — 載入後可逐張確認圖片，點擊排除不需要的圖

## 快速開始

### 使用方式

1. 用 Chrome 瀏覽器打開 `docx-images-to-ppt-offline.html`
2. 拖放一或多個 `.docx` 檔案（或點擊選擇，支援 Ctrl 多選）
3. 確認預覽中的圖片，可點擊排除不需要的
4. 點擊「生成 PPT」，瀏覽器自動下載 `.pptx` 檔案

就這麼簡單，不需要安裝任何東西。

### 環境需求

- **瀏覽器**：Chrome 80+ / Edge 80+ / Firefox 78+
- **網路**：不需要（完全離線）
- **Python / Node.js**：不需要

## 批量排除設定

每個 Word 檔的前幾張圖片通常是頁首頁尾的 logo 或浮水印。你可以設定自動排除：

| 格式 | 效果 |
|------|------|
| `1-6` | 排除第 1~6 張 |
| `1,3,5` | 排除第 1、3、5 張 |
| `1-6, 8, 10-12` | 混合寫法 |

- 預設排除每個檔案的 **第 1~6 張**
- 載入時自動生效
- 也可在預覽區改範圍後點「套用排除規則」重新套用

## 生成的 PPT 結構

```
📄 封面（SEM-EDS 圖片彙整 + 統計資訊）
  │
  ├── 📄 檔案 1 標題頁（檔名）
  ├── 📄 檔案 1 圖片總覽（所有圖片排成網格，一張 slide）
  │     └── 若圖片太多會自動分頁（1/2, 2/2...）
  │
  ├── 📄 檔案 2 標題頁
  ├── 📄 檔案 2 圖片總覽
  │
  └── ...（依此類推）
```

### 自動網格排版邏輯

根據圖片數量自動選擇最佳網格：

| 圖片數 | 網格 |
|--------|------|
| 1 | 滿版 |
| 2~3 | 橫排 |
| 4 | 2×2 |
| 5~6 | 3×2 |
| 7~9 | 3×3 |
| 10~12 | 4×3 |
| 13~16 | 4×4 |
| 17~20 | 5×4 |
| 21~25 | 5×5 |
| 26+ | 自動計算（基於 16:9 比例） |

超出單頁容量時自動分頁。

## 技術原理

### .docx 內部結構

`.docx` 本質上是 ZIP 壓縮檔，圖片通常存放在 `word/media/` 目錄下。本工具的處理流程：

1. **ZIP 解壓縮** — 使用 JSZip 讀取 .docx 內部結構
2. **圖片掃描** — 掃描整個 ZIP 中所有圖片格式檔案（png/jpg/gif/bmp/emf/wmf/svg 等）
3. **順序解析** — 解析 `document.xml` 中的 `<a:blip>` 和 `<v:imagedata>` 元素，追蹤 `r:embed` 引用
4. **關聯映射** — 對照 `_rels/document.xml.rels` 映射關係 ID 到實際檔案路徑
5. **智能合併** — XML 解析結果不足時，自動補充媒體資料夾掃描結果
6. **網格排版** — 根據圖片數量自動計算最佳 cols × rows 網格，等比縮放填入 slide

### 內嵌函式庫

| 函式庫 | 版本 | 用途 |
|--------|------|------|
| [JSZip](https://github.com/Stuk/jszip) | 3.10.1 | 解壓縮 .docx (ZIP) |
| [PptxGenJS](https://github.com/gitbrent/PptxGenJS) | 3.12.0 | 生成 .pptx 檔案 |

兩個函式庫已直接內嵌進 HTML，無需從 CDN 載入。

## 檔案說明

```
docx-images-to-ppt-offline.html   ← 唯一需要的檔案，完全獨立運行
```

就一個 HTML 檔案，沒有其他依賴。

## 已知限制

- **EMF/WMF 格式**：Office 舊式向量圖格式在 PPT 中可能無法完美顯示，建議在 Word 中先轉為 PNG
- **超大圖片**：單張超過 5MB 的圖片可能導致生成速度較慢
- **圖表/SmartArt**：Word 中的圖表和 SmartArt 通常以多層結構嵌入，提取時可能只會抓到其中一層
- **頁首頁尾圖片**：可能出現在 media 資料夾中，請使用批量排除功能過濾

## 開發與維護

### 修改工具

直接編輯 `docx-images-to-ppt-offline.html` 即可。檔案結構：

```
<html>
  <head>
    <script> JSZip 內嵌 </script>
    <script> PptxGenJS 內嵌 </script>
    <style> 所有 CSS </style>
  </head>
  <body>
    UI 結構
    <script> 應用邏輯 </script>
  </body>
</html>
```

### 關鍵函式

| 函式 | 說明 |
|------|------|
| `handleFiles(files)` | 主流程：解析 ZIP → 掃描圖片 → 提取 → 預覽 |
| `parseExcludeRange(str)` | 解析排除範圍字串（如 `1-6, 8`） |
| `pickGrid(n)` | 根據圖片數量計算最佳網格 cols × rows |
| `generatePPT()` | 生成 PPT：封面 → 各檔標題頁 → 圖片網格頁 |
| `renderPreview()` | 渲染預覽區（含點擊排除功能） |
| `applyBatchExclude()` | 套用/重新套用批量排除規則 |

### 更新內嵌函式庫

如需更新 JSZip 或 PptxGenJS 版本：

```bash
# 下載新版
curl -sL "https://unpkg.com/jszip@latest/dist/jszip.min.js" -o jszip.min.js
curl -sL "https://unpkg.com/pptxgenjs@latest/dist/pptxgen.bundle.js" -o pptxgen.bundle.js

# 替換 HTML 中的 <script> 區塊
# JSZip → 第一個 <script> 標籤內容
# PptxGenJS → 第二個 <script> 標籤內容
```

## License

MIT License
