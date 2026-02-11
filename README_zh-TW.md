# MermaidPreview

EmEditor 原生外掛程式 — 在可停靠的側邊欄中即時預覽 Mermaid 圖表與 Markdown 內容。

![Version](https://img.shields.io/badge/version-1.1.0-blue)
![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-lightgrey)
![C++17](https://img.shields.io/badge/C%2B%2B-17-orange)
![License](https://img.shields.io/badge/license-MIT-green)

## 功能特色

- **即時 Mermaid 圖表** — 支援流程圖、序列圖、類別圖、甘特圖等所有 Mermaid 支援的圖表類型，即時渲染
- **原生 Markdown 渲染** — 內建 C++ 解析器處理標題、清單、表格、程式碼區塊、連結、圖片、粗體/斜體、刪除線、任務清單和引用區塊，零外部依賴
- **雙向捲動同步** — 編輯器與預覽面板的捲動位置雙向同步
- **行內編輯** — 雙擊預覽面板中的文字區塊可直接編輯，修改內容會寫回編輯器
- **亮色 / 暗色主題** — 自動偵測 EmEditor 主題，亦可透過右鍵選單手動切換
- **點擊導航** — 點擊 Mermaid 圖表可跳轉至編輯器中對應的原始碼行
- **SVG 平移與縮放** — Ctrl+點擊圖表啟用縮放模式；Ctrl+滾輪可直接縮放；拖曳平移
- **自動開啟/關閉** — 偵測到含有 ` ```mermaid ` 區塊的 Markdown 檔案時自動開啟；切換到非 Mermaid 檔案或關閉標籤時自動關閉
- **快速重開（Parking Window）** — 關閉預覽時 WebView2 會停泊至隱藏視窗而非銷毀；重開僅需 ~10ms（原需 ~800-1500ms）
- **可停靠面板** — 支援左、右、上、下停靠，位置跨工作階段保存
- **安全強化** — URL scheme 驗證封鎖 `javascript:`/`data:`/`vbscript:` 連結；WebView2 host object 注入已停用；行號輸入已驗證
- **Bun 預渲染**（選用）— 安裝 [Bun](https://bun.sh) 後，Mermaid 圖表可在伺服器端預先渲染為 SVG，顯示更快

## 架構

```
EmEditor
 └── MermaidPreview.dll (C++17 / MSVC)
      ├── MermaidPreview   — 外掛生命週期、Custom Bar、事件處理
      ├── MarkdownParser   — C++ 原生 Markdown → HTML 轉換器
      ├── WebView2Manager  — WebView2 初始化、JS 互動、HTML 外殼
      └── BunRenderer      — 選用的伺服器端 Mermaid → SVG（透過 Bun/jsdom）
           └── bun-renderer/renderer.ts
```

### 渲染流程

```
編輯器文字
  → MarkdownParser::ConvertToHtml()      (C++ 原生，~1ms)
  → BunRenderer::RenderBlocks()          (選用，將 mermaid 預留位置替換為 SVG)
  → WebView2 renderContent()             (顯示 HTML + 以 JS 渲染剩餘的 mermaid)
```

### 效能優化

| 優化項目 | 說明 | 節省時間 |
|---|---|---|
| **HTML 快取** | `preview.html` 以版本標記快取於磁碟，未變更時跳過重建 | ~10-20ms |
| **Async mermaid.js** | mermaid.min.js (~2MB) 非同步載入，Markdown 文字立即顯示 | ~200-500ms |
| **內容預取** | 文件解析與 WebView2 初始化平行執行 | ~10-50ms |
| **Parking Window** | 關閉時 WebView2 停泊至隱藏視窗而非銷毀；重開跳過完整初始化 | ~800-1500ms |

## 系統需求

| 元件 | 版本 | 備註 |
|---|---|---|
| **EmEditor** | Professional v20+ | 需要 Custom Bar API 支援 |
| **Windows** | 10 / 11 | 需要 WebView2 Runtime |
| **WebView2 Runtime** | 最新版 | Windows 11 已內建；[Windows 10 下載](https://developer.microsoft.com/en-us/microsoft-edge/webview2/) |
| **Bun** | v1.0+ | *選用* — 用於伺服器端 Mermaid 預渲染 |

## 安裝方式

### 從 Release 安裝

1. 從 [Releases](https://github.com/CHC1978/MermaidPreview/releases) 頁面下載 `MermaidPreview.dll`
2. 複製到 EmEditor 的 `PlugIns` 目錄：
   ```
   %APPDATA%\Emurasoft\EmEditor\PlugIns\
   ```
   或 `EmEditor.exe` 旁邊的 `PlugIns` 資料夾
3. 重新啟動 EmEditor
4. 工具列中會出現 **Mermaid Preview** 按鈕

### 選用：Bun 渲染器

如果想要透過伺服器端預渲染加速 Mermaid 圖表：

```bash
# 安裝 Bun（如尚未安裝）
powershell -c "irm bun.sh/install.ps1 | iex"

# 安裝相依套件
cd bun-renderer
bun install
```

## 從原始碼建置

### 前置條件

- **MSVC 2022**（Visual Studio 2022，需安裝 C++ 桌面開發工作負載）
- **CMake** 3.24+
- **Ninja**（建議使用）或 MSBuild

### 建置步驟

```bash
# 組態設定（WebView2 SDK 會自動下載）
cmake --preset x64-release

# 建置
cmake --build build

# 輸出：build/MermaidPreview.dll
```

### Debug 建置

```bash
cmake --preset x64-debug
cmake --build build-debug
```

## 使用方式

1. 在 EmEditor 中開啟 Markdown 檔案（`.md`、`.markdown`）
2. 點擊工具列的 **Mermaid Preview** 按鈕（或外掛偵測到 mermaid 區塊時自動開啟）
3. 側邊欄出現預覽面板，顯示渲染後的 Markdown 與 Mermaid 圖表
4. 編輯文件 — 預覽會以 debounce 機制即時更新

### 鍵盤與滑鼠操作

| 動作 | 效果 |
|---|---|
| **編輯文字** | 編輯器中的變更會在短暫延遲後反映到預覽 |
| **雙擊**預覽文字 | 啟用行內編輯；按 Enter 儲存，Escape 取消 |
| **右鍵**預覽面板 | 開啟右鍵選單（主題切換、縮放控制） |
| **點擊**圖表 | 跳轉至編輯器中對應的原始碼行 |
| **Ctrl+點擊**圖表 | 啟用該圖表的縮放模式 |
| **Ctrl+滾輪**於圖表上 | 直接縮放（無需先啟用縮放模式） |
| **滾輪**於啟用的圖表上 | 縮放 |
| **拖曳**啟用的圖表 | 平移圖表 |

## 專案結構

```
MermaidPreview/
├── CMakeLists.txt          # 建置設定
├── CMakePresets.json        # MSVC 2022 預設組態
├── exports.def              # DLL 匯出定義
├── include/
│   ├── etlframe.h           # EmEditor Template Library
│   ├── plugin.h             # EmEditor Plugin SDK
│   └── resource.h           # 資源 ID
├── src/
│   ├── DllMain.cpp          # DLL 進入點
│   ├── MermaidPreview.cpp   # 外掛主要邏輯
│   ├── MermaidPreview.h
│   ├── MarkdownParser.cpp   # C++ Markdown → HTML
│   ├── MarkdownParser.h
│   ├── WebView2Manager.cpp  # WebView2 生命週期與 JS
│   ├── WebView2Manager.h
│   ├── BunRenderer.cpp      # Bun IPC 用於 mermaid SVG
│   └── BunRenderer.h
├── resources/
│   ├── MermaidPreview.rc    # 資源腳本
│   ├── icon_16.bmp          # 16x16 工具列圖示
│   ├── icon_24.bmp          # 24x24 工具列圖示
│   └── web/
│       └── mermaid.min.js   # 內嵌的 mermaid.js (~2MB)
└── bun-renderer/
    ├── package.json
    └── renderer.ts          # Bun 端 mermaid 渲染器
```

## 運作原理

1. **外掛載入** → 註冊為 EmEditor 外掛程式，新增工具列按鈕
2. **使用者點擊按鈕** → 建立可停靠的 Custom Bar 及子視窗
3. **WebView2 初始化** → 載入本地 HTML 外殼（`preview.html`），包含 CSS 與 JavaScript
4. **Markdown 解析** → `MarkdownParser` 將編輯器內容轉換為帶有行號追蹤的 HTML（`data-line-start` / `data-line-end` 屬性）
5. **Mermaid 渲染** → Mermaid 程式碼區塊轉為 `<div class="mermaid-container">` 預留位置；可選擇由 Bun 預渲染，否則由 mermaid.js 在客戶端渲染
6. **即時更新** → `EVENT_MODIFIED` 觸發 debounce 重新渲染；`EVENT_SCROLL` 觸發捲動同步
7. **雙向同步** — 行號屬性實現編輯器與預覽之間的精確捲動對應

## 授權條款

MIT

## 致謝

- [Mermaid.js](https://mermaid.js.org/) — 圖表渲染引擎
- [Microsoft WebView2](https://developer.microsoft.com/en-us/microsoft-edge/webview2/) — 內嵌瀏覽器控制項
- [Bun](https://bun.sh/) — JavaScript 執行環境，用於伺服器端渲染
- [EmEditor](https://www.emeditor.com/) — 文字編輯器與外掛 SDK
