# MermaidPreview

A native EmEditor plug-in that renders live Mermaid diagrams and Markdown previews in a dockable sidebar panel.

![Version](https://img.shields.io/badge/version-1.2.0-blue)
![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-lightgrey)
![C++17](https://img.shields.io/badge/C%2B%2B-17-orange)
![Mermaid](https://img.shields.io/badge/mermaid-11.14.0-ff3670)
![License](https://img.shields.io/badge/license-MIT-green)

## Features

- **Live Mermaid Diagrams** — Flowcharts, sequence diagrams, class diagrams, Gantt charts, state diagrams, ER diagrams, mindmaps, timelines, GitGraph, architecture diagrams, **Wardley Maps (beta)**, **TreeView**, and every other Mermaid-supported type, rendered in real time
- **Mermaid 11.14 Neo Look** — Opt-in modern visual style with drop-shadows for flowcharts/class/state/sequence/mindmap/GitGraph; just add `%%{init: {'look': 'neo'}}%%` at the top of any diagram
- **Native Markdown Rendering** — C++ built-in parser handles headings, lists, tables, code blocks, links, images, bold/italic, strikethrough, task lists, and blockquotes with zero external dependencies
- **Bidirectional Scroll Sync** — Editor and preview scroll positions stay synchronized in both directions
- **Inline Editing** — Double-click any text block in the preview to edit it directly; changes are written back to the editor
- **Light / Dark Theme** — Auto-detects EmEditor's theme; manual override available via right-click context menu
- **Click-to-Navigate** — Click any Mermaid diagram to jump the editor cursor to the corresponding source line
- **SVG Pan & Zoom** — Ctrl+Click a diagram to activate zoom mode; Ctrl+Scroll to zoom directly; drag to pan
- **Auto Open/Close** — Automatically opens when a Markdown file containing ` ```mermaid ` blocks is detected; closes when switching to non-Mermaid files or closing the tab
- **Fast Reopen (Parking Window)** — When the preview is closed, WebView2 is parked in a hidden window instead of destroyed; reopening takes ~10ms instead of ~800-1500ms
- **Non-blocking Bun Render** — Server-side mermaid rendering runs on a background thread with a 15 s safety timeout; the editor UI never freezes even when Bun stalls
- **Dockable Panel** — Supports left, right, top, or bottom docking; position is persisted across sessions
- **Security Hardened** — Path-traversal guarded relative-link opener (`../../etc/passwd` is rejected); URL scheme deny-list blocks `javascript:` / `data:` / `vbscript:` / `blob:` / `file:`; WebMessage dispatch uses exact `type` matching to prevent type confusion; embedded null/control bytes rejected; OOM caps on every Bun pipe buffer (20 MB) and per-block SVG (10 MB); UTF-16 surrogate pairs preserved through URL encoding
- **Bun Pre-rendering** (Optional) — If [Bun](https://bun.sh) v1.3+ is installed, Mermaid diagrams are pre-rendered server-side for faster display

## Architecture

```
EmEditor
 └── MermaidPreview.dll (C++17 / MSVC)
      ├── MermaidPreview   — Plugin lifecycle, Custom Bar, async render coordination
      ├── MarkdownParser   — C++ native Markdown → HTML converter
      ├── WebView2Manager  — WebView2 initialization, JS interop, HTML shell
      └── BunRenderer      — Optional server-side Mermaid → SVG via Bun/jsdom
           └── bun-renderer/renderer.ts
```

### Rendering Pipeline

```
Editor Text
  → MarkdownParser::ConvertToHtml()      (C++ native, ~1 ms)
  → WebView2 renderContent(placeholders) (UI shows text immediately)
  → BunRenderer::RenderBlocks()          (background thread, 50–500 ms typical)
  → IDT_BUN_POLL fires when future is ready
  → SpliceSvgIntoHtml() + WebView2 renderContent(SVG)
  └─ Fallback: if Bun is unavailable, mermaid.js renders placeholders client-side.
```

### Performance Optimizations

| Optimization | Description | Savings |
|---|---|---|
| **HTML Cache** | `preview.html` is cached on disk with a version tag; skips rebuild when unchanged | ~10–20 ms |
| **Async mermaid.js** | mermaid.min.js (~3.1 MB) loads asynchronously; Markdown text appears immediately | ~200–500 ms |
| **Content Pre-fetch** | Document parsing runs in parallel with WebView2 initialization | ~10–50 ms |
| **Parking Window** | WebView2 is reparented to a hidden window on close instead of destroyed; reopen skips full init | ~800–1500 ms |
| **Background Bun render** | `RenderBlocks` runs on `std::async` worker; UI thread polls via 40 ms timer; 15 s safety cap, dirty-flag re-trigger on rapid edits | UI never blocks |

## Requirements

| Component | Version | Notes |
|---|---|---|
| **EmEditor** | Professional v20+ | Required for Custom Bar API |
| **Windows** | 10 / 11 | WebView2 Runtime required |
| **WebView2 Runtime** | Latest | Pre-installed on Windows 11; [download for Windows 10](https://developer.microsoft.com/en-us/microsoft-edge/webview2/) |
| **Mermaid** | 11.14.0 | Embedded in DLL as RCDATA — no separate install |
| **Bun** | v1.3+ | *Optional* — for server-side Mermaid pre-rendering |

## Installation

### From Release

1. Download `MermaidPreview.dll` from the [Releases](https://github.com/CHC1978/MermaidPreview/releases) page
2. Copy it to EmEditor's `PlugIns` directory:
   ```
   %APPDATA%\Emurasoft\EmEditor\PlugIns\
   ```
   or the `PlugIns` folder next to `EmEditor.exe`
3. Restart EmEditor
4. The **Mermaid Preview** button appears in the toolbar

### Optional: Bun Renderer

If you want faster Mermaid rendering via server-side pre-rendering:

```bash
# Install Bun (if not installed)
powershell -c "irm bun.sh/install.ps1 | iex"

# Install dependencies
cd bun-renderer
bun install
```

## Building from Source

### Prerequisites

- **MSVC 2022** (Visual Studio 2022 with C++ Desktop workload)
- **CMake** 3.24+
- **Ninja** (recommended) or MSBuild

### Build Steps

```bash
# Configure (WebView2 SDK is fetched automatically)
cmake --preset x64-release

# Build
cmake --build build

# Output: build/MermaidPreview.dll
```

### Debug Build

```bash
cmake --preset x64-debug
cmake --build build-debug
```

## Usage

1. Open a Markdown file (`.md`, `.markdown`) in EmEditor
2. Click the **Mermaid Preview** toolbar button (or the plugin auto-opens if mermaid blocks are detected)
3. The preview panel appears in the sidebar showing rendered Markdown and Mermaid diagrams
4. Edit your document — the preview updates in real-time with debouncing

### Keyboard & Mouse

| Action | Effect |
|---|---|
| **Edit text** | Changes in the editor are reflected in the preview after a short delay |
| **Double-click** preview text | Enables inline editing; press Enter to save, Escape to cancel |
| **Right-click** preview | Opens context menu (theme toggle, zoom controls, font size) |
| **Click** a diagram | Jumps the editor cursor to the corresponding source line |
| **Ctrl+Click** a diagram | Activates zoom mode for that diagram |
| **Ctrl+Scroll wheel** on a diagram | Zoom in/out directly (no need to activate zoom first) |
| **Scroll wheel** on active diagram | Zoom in/out |
| **Drag** an active diagram | Pan the diagram |

### Enabling the Neo Look

Add a single `%%{init}%%` directive at the top of any diagram:

````markdown
```mermaid
%%{init: {'look': 'neo'}}%%
flowchart TD
    A[Start] --> B{Decide}
    B -->|Yes| C[Done]
    B -->|No|  D[Retry]
```
````

Valid `look` values: `classic` (default), `neo`, `handDrawn`.

## Project Structure

```
MermaidPreview/
├── CMakeLists.txt          # Build configuration
├── CMakePresets.json        # MSVC 2022 presets
├── exports.def              # DLL export definitions
├── include/
│   ├── etlframe.h           # EmEditor Template Library
│   ├── plugin.h             # EmEditor Plugin SDK
│   └── resource.h           # Resource IDs
├── src/
│   ├── DllMain.cpp          # DLL entry point
│   ├── MermaidPreview.cpp   # Plugin main logic, async render coordinator
│   ├── MermaidPreview.h
│   ├── MarkdownParser.cpp   # C++ Markdown → HTML
│   ├── MarkdownParser.h
│   ├── WebView2Manager.cpp  # WebView2 lifecycle & JS
│   ├── WebView2Manager.h
│   ├── BunRenderer.cpp      # Bun IPC for mermaid SVG
│   └── BunRenderer.h
├── resources/
│   ├── MermaidPreview.rc    # Resource script
│   ├── icon_16.bmp          # 16x16 toolbar icon
│   ├── icon_24.bmp          # 24x24 toolbar icon
│   └── web/
│       └── mermaid.min.js   # Embedded mermaid 11.14.0 (~3.1 MB)
└── bun-renderer/
    ├── package.json         # mermaid 11.14, jsdom
    └── renderer.ts          # Bun-based mermaid renderer
```

## How It Works

1. **Plugin loads** → Registers as an EmEditor plug-in with toolbar button
2. **User clicks button** → Creates a dockable Custom Bar with a child window
3. **WebView2 initializes** → Loads a local HTML shell containing CSS and a single `_mmdInit(theme, look)` helper that drives all `mermaid.initialize` call sites
4. **Markdown parsing** → `MarkdownParser` converts editor content to HTML with line-number tracking (`data-line-start` / `data-line-end` attributes)
5. **Mermaid rendering** — async path:
   - Mermaid code blocks become `<div class="mermaid-container">` placeholders
   - Placeholder HTML is shipped to WebView2 immediately so the user sees text
   - `BunRenderer::RenderBlocks` runs on a worker thread; `IDT_BUN_POLL` fires every 40 ms on the UI thread
   - When the future resolves, SVGs are spliced into the cached HTML and rendered; if rendering takes longer than 15 s the worker is abandoned and the WebView's client-side mermaid.js takes over
6. **Live updates** — `EVENT_MODIFIED` triggers debounced re-render; `EVENT_SCROLL` triggers scroll sync; a `m_renderDirty` flag re-runs the pipeline if the document changed during a Bun render
7. **Bidirectional sync** — Line-number attributes enable precise scroll mapping between editor and preview

## Security & Robustness Highlights

- **Path traversal**: relative-path file links are resolved with `GetFullPathNameW`, then bounded against the current document's directory via case-insensitive `_wcsnicmp` comparison; `[evil](../../Windows/System32/notepad.exe)` is rejected
- **URL scheme deny-list**: `javascript:`, `vbscript:`, `data:`, `blob:`, `file:`, `ms-appx:`, `ms-its:`, `mhtml:`, `ms-msdt:`, `ms-help:`, plus null/control-byte rejection so `java\0script:` cannot smuggle past
- **WebMessage type confusion**: a `"type"` field is extracted exactly once and dispatched on equality (`msgType == L"theme"`), preventing payloads that merely contain the string `"theme"` from hijacking the theme handler
- **Resource caps**: 20 MB on the Bun stdout read buffer (kills runaway processes), 10 MB per spliced SVG, 15 s pipe timeout, recursive markdown depth ≤ 20
- **Memory safety**: `std::async` worker captures `shared_ptr<BunRenderer>` so the renderer outlives any in-flight render; `CloseCustomBar` and `~CMermaidFrame` detach the future to a graveyard thread so the destructor cannot stall the UI close path or DLL unload

## License

MIT

## Credits

- [Mermaid.js](https://mermaid.js.org/) — Diagram rendering engine (v11.14.0)
- [Microsoft WebView2](https://developer.microsoft.com/en-us/microsoft-edge/webview2/) — Embedded browser control
- [Bun](https://bun.sh/) — JavaScript runtime for server-side rendering
- [EmEditor](https://www.emeditor.com/) — Text editor and plug-in SDK
