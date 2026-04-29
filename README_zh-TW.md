# MermaidPreview

EmEditor 原生外掛程式 — 在可停靠的側邊欄中即時預覽 Mermaid 圖表與 Markdown 內容。

![Version](https://img.shields.io/badge/version-1.2.0-blue)
![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-lightgrey)
![C++17](https://img.shields.io/badge/C%2B%2B-17-orange)
![Mermaid](https://img.shields.io/badge/mermaid-11.14.0-ff3670)
![License](https://img.shields.io/badge/license-MIT-green)

## 功能特色

- **即時 Mermaid 圖表** — 支援流程圖、序列圖、類別圖、甘特圖、狀態圖、ER 圖、心智圖、時間軸、GitGraph、架構圖、**Wardley Maps（beta）**、**TreeView** 等所有 Mermaid 圖表類型，即時渲染
- **Mermaid 11.14 Neo Look** — 可選的現代視覺風格，flowchart/class/state/sequence/mindmap/GitGraph 都會套上 drop-shadow；只要在圖表開頭加入 `%%{init: {'look': 'neo'}}%%` 即可啟用
- **原生 Markdown 渲染** — 內建 C++ 解析器處理標題、清單、表格、程式碼區塊、連結、圖片、粗體/斜體、刪除線、任務清單和引用區塊，零外部依賴
- **雙向捲動同步** — 編輯器與預覽面板的捲動位置雙向同步
- **行內編輯** — 雙擊預覽面板中的文字區塊可直接編輯，修改內容會寫回編輯器
- **亮色 / 暗色主題** — 自動偵測 EmEditor 主題，亦可透過右鍵選單手動切換
- **點擊導航** — 點擊 Mermaid 圖表可跳轉至編輯器中對應的原始碼行
- **SVG 平移與縮放** — Ctrl+點擊圖表啟用縮放模式；Ctrl+滾輪可直接縮放；拖曳平移
- **自動開啟/關閉** — 偵測到含有 ` ```mermaid ` 區塊的 Markdown 檔案時自動開啟；切換到非 Mermaid 檔案或關閉標籤時自動關閉
- **快速重開（Parking Window）** — 關閉預覽時 WebView2 會停泊至隱藏視窗而非銷毀；重開僅需 ~10ms（原需 ~800-1500ms）
- **非阻塞 Bun 渲染** — 伺服器端 mermaid 渲染移到背景緒，並設 15 秒安全 timeout；即使 Bun 卡住，編輯器 UI 也絕不凍結
- **可停靠面板** — 支援左、右、上、下停靠，位置跨工作階段保存
- **安全強化** — 相對路徑連結加上路徑遍歷防護（`../../etc/passwd` 一律拒絕）；URL scheme 黑名單封鎖 `javascript:` / `data:` / `vbscript:` / `blob:` / `file:`；WebMessage 用精確 `type` 比對防止 type confusion；拒絕嵌入 null / 控制位元組；Bun pipe buffer 20 MB 上限、單一 SVG 10 MB 上限、UTF-16 surrogate pair 完整保留
- **Bun 預渲染**（選用）— 安裝 [Bun](https://bun.sh) v1.3+ 後，Mermaid 圖表可在伺服器端預先渲染為 SVG，顯示更快

## 架構

```
EmEditor
 └── MermaidPreview.dll (C++17 / MSVC)
      ├── MermaidPreview   — 外掛生命週期、Custom Bar、非同步渲染協調
      ├── MarkdownParser   — C++ 原生 Markdown → HTML 轉換器
      ├── WebView2Manager  — WebView2 初始化、JS 互動、HTML 外殼
      └── BunRenderer      — 選用的伺服器端 Mermaid → SVG（透過 Bun/jsdom）
           └── bun-renderer/renderer.ts
```

### 渲染流程

```
編輯器文字
  → MarkdownParser::ConvertToHtml()      (C++ 原生，~1 ms)
  → WebView2 renderContent(預留位置)     (UI 立即顯示文字)
  → BunRenderer::RenderBlocks()          (背景緒，典型 50–500 ms)
  → IDT_BUN_POLL 偵測 future 完成
  → SpliceSvgIntoHtml() + WebView2 renderContent(SVG)
  └─ Fallback: 若 Bun 不可用，由 mermaid.js 在客戶端渲染預留位置
```

### 效能優化

| 優化項目 | 說明 | 節省時間 |
|---|---|---|
| **HTML 快取** | `preview.html` 以版本標記快取於磁碟，未變更時跳過重建 | ~10–20 ms |
| **Async mermaid.js** | mermaid.min.js (~3.1 MB) 非同步載入，Markdown 文字立即顯示 | ~200–500 ms |
| **內容預取** | 文件解析與 WebView2 初始化平行執行 | ~10–50 ms |
| **Parking Window** | 關閉時 WebView2 停泊至隱藏視窗而非銷毀；重開跳過完整初始化 | ~800–1500 ms |
| **背景 Bun 渲染** | `RenderBlocks` 跑在 `std::async` worker；UI 緒以 40 ms 計時器輪詢；15 秒安全 cap、dirty flag 串連快速編輯 | UI 永不阻塞 |

## 系統需求

| 元件 | 版本 | 備註 |
|---|---|---|
| **EmEditor** | Professional v20+ | 需要 Custom Bar API 支援 |
| **Windows** | 10 / 11 | 需要 WebView2 Runtime |
| **WebView2 Runtime** | 最新版 | Windows 11 已內建；[Windows 10 下載](https://developer.microsoft.com/en-us/microsoft-edge/webview2/) |
| **Mermaid** | 11.14.0 | 已內嵌於 DLL（RCDATA），無需額外安裝 |
| **Bun** | v1.3+ | *選用* — 用於伺服器端 Mermaid 預渲染 |

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
| **右鍵**預覽面板 | 開啟右鍵選單（主題切換、縮放控制、字型大小） |
| **點擊**圖表 | 跳轉至編輯器中對應的原始碼行 |
| **Ctrl+點擊**圖表 | 啟用該圖表的縮放模式 |
| **Ctrl+滾輪**於圖表上 | 直接縮放（無需先啟用縮放模式） |
| **滾輪**於啟用的圖表上 | 縮放 |
| **拖曳**啟用的圖表 | 平移圖表 |

### 啟用 Neo Look

在任一圖表的開頭加上一行 `%%{init}%%` directive 即可：

````markdown
```mermaid
%%{init: {'look': 'neo'}}%%
flowchart TD
    A[Start] --> B{Decide}
    B -->|Yes| C[Done]
    B -->|No|  D[Retry]
```
````

`look` 可用值：`classic`（預設）、`neo`、`handDrawn`。

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
│   ├── MermaidPreview.cpp   # 外掛主要邏輯、非同步渲染協調器
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
│       └── mermaid.min.js   # 內嵌 mermaid 11.14.0 (~3.1 MB)
└── bun-renderer/
    ├── package.json         # mermaid 11.14、jsdom
    └── renderer.ts          # Bun 端 mermaid 渲染器
```

## 運作原理

1. **外掛載入** → 註冊為 EmEditor 外掛程式，新增工具列按鈕
2. **使用者點擊按鈕** → 建立可停靠的 Custom Bar 及子視窗
3. **WebView2 初始化** → 載入本地 HTML 外殼，內含 CSS 與單一 `_mmdInit(theme, look)` helper 統一驅動所有 `mermaid.initialize` 呼叫站
4. **Markdown 解析** → `MarkdownParser` 將編輯器內容轉換為帶有行號追蹤的 HTML（`data-line-start` / `data-line-end` 屬性）
5. **Mermaid 渲染（非同步路徑）：**
   - Mermaid 程式碼區塊轉為 `<div class="mermaid-container">` 預留位置
   - 預留位置 HTML 立即送往 WebView2，使用者馬上看到文字
   - `BunRenderer::RenderBlocks` 跑在 worker thread；UI 緒以 `IDT_BUN_POLL`（40 ms）輪詢
   - future 完成時，將 SVG 拼接回快取 HTML 並重新渲染；超過 15 秒則放棄背景 worker，由 WebView 內 mermaid.js 接手
6. **即時更新** — `EVENT_MODIFIED` 觸發 debounce 重新渲染；`EVENT_SCROLL` 觸發捲動同步；`m_renderDirty` 旗標會在 Bun 渲染期間若文件再次修改時自動重跑
7. **雙向同步** — 行號屬性實現編輯器與預覽之間的精確捲動對應

## 安全性與穩健度重點

- **路徑遍歷防護**：相對路徑檔案連結先用 `GetFullPathNameW` 正規化，再以大小寫不敏感的 `_wcsnicmp` 比對是否仍在當前文件目錄內；`[evil](../../Windows/System32/notepad.exe)` 直接拒絕
- **URL scheme 黑名單**：`javascript:`、`vbscript:`、`data:`、`blob:`、`file:`、`ms-appx:`、`ms-its:`、`mhtml:`、`ms-msdt:`、`ms-help:`，且拒絕嵌入 null / 控制位元組，避免 `java\0script:` 這類繞過手法
- **WebMessage type confusion**：`"type"` 欄位精確抽取一次後以 equality 比對（`msgType == L"theme"`），含有 `"theme"` 字串的其他訊息無法劫持 theme handler
- **資源上限**：Bun stdout buffer 20 MB（觸及即殺掉跑飛的進程）、單一 SVG 拼接 10 MB、pipe timeout 15 秒、Markdown 遞迴深度上限 20
- **記憶體安全**：`std::async` worker 透過 `shared_ptr<BunRenderer>` 捕獲 renderer，確保渲染進行中 renderer 不會被銷毀；`CloseCustomBar` 與 `~CMermaidFrame` 將 in-flight future 移交至 graveyard thread，避免 future destructor 阻塞 UI 關閉路徑或 DLL unload

## 授權條款

MIT

## 致謝

- [Mermaid.js](https://mermaid.js.org/) — 圖表渲染引擎（v11.14.0）
- [Microsoft WebView2](https://developer.microsoft.com/en-us/microsoft-edge/webview2/) — 內嵌瀏覽器控制項
- [Bun](https://bun.sh/) — JavaScript 執行環境，用於伺服器端渲染
- [EmEditor](https://www.emeditor.com/) — 文字編輯器與外掛 SDK
