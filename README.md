# MermaidPreview

A native EmEditor plug-in that renders live Mermaid diagrams and Markdown previews in a dockable sidebar panel.

![Version](https://img.shields.io/badge/version-1.0.0-blue)
![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-lightgrey)
![C++17](https://img.shields.io/badge/C%2B%2B-17-orange)
![License](https://img.shields.io/badge/license-MIT-green)

## Features

- **Live Mermaid Diagrams** — Flowcharts, sequence diagrams, class diagrams, Gantt charts, and all Mermaid-supported diagram types rendered in real-time
- **Native Markdown Rendering** — C++ built-in parser handles headings, lists, tables, code blocks, links, images, bold/italic, strikethrough, task lists, and blockquotes with zero external dependencies
- **Bidirectional Scroll Sync** — Editor and preview scroll positions stay synchronized in both directions
- **Inline Editing** — Double-click any text block in the preview to edit it directly; changes are written back to the editor
- **Light / Dark Theme** — Auto-detects EmEditor's theme; manual override available via right-click context menu
- **SVG Pan & Zoom** — Click a diagram to activate zoom mode, then scroll to zoom in/out; drag to pan
- **Auto Open/Close** — Automatically opens when a Markdown file containing ` ```mermaid ` blocks is detected; closes when switching to non-Mermaid files
- **Dockable Panel** — Supports left, right, top, or bottom docking; position is persisted across sessions
- **Bun Pre-rendering** (Optional) — If [Bun](https://bun.sh) is installed, Mermaid diagrams are pre-rendered server-side for faster display

## Architecture

```
EmEditor
 └── MermaidPreview.dll (C++17 / MSVC)
      ├── MermaidPreview   — Plugin lifecycle, Custom Bar, event handling
      ├── MarkdownParser   — C++ native Markdown → HTML converter
      ├── WebView2Manager  — WebView2 initialization, JS interop, HTML shell
      └── BunRenderer      — Optional server-side Mermaid → SVG via Bun/jsdom
           └── bun-renderer/renderer.ts
```

### Rendering Pipeline

```
Editor Text
  → MarkdownParser::ConvertToHtml()      (C++ native, ~1ms)
  → BunRenderer::RenderBlocks()          (optional, replaces mermaid placeholders with SVG)
  → WebView2 renderContent()             (displays HTML + renders remaining mermaid via JS)
```

### Performance Optimizations

| Optimization | Description | Savings |
|---|---|---|
| **HTML Cache** | `preview.html` is cached on disk with a version tag; skips rebuild when unchanged | ~10-20ms |
| **Async mermaid.js** | mermaid.min.js (~2MB) loads asynchronously; Markdown text appears immediately | ~200-500ms |
| **Content Pre-fetch** | Document parsing runs in parallel with WebView2 initialization | ~10-50ms |

## Requirements

| Component | Version | Notes |
|---|---|---|
| **EmEditor** | Professional v20+ | Required for Custom Bar API |
| **Windows** | 10 / 11 | WebView2 Runtime required |
| **WebView2 Runtime** | Latest | Pre-installed on Windows 11; [download for Windows 10](https://developer.microsoft.com/en-us/microsoft-edge/webview2/) |
| **Bun** | v1.0+ | *Optional* — for server-side Mermaid pre-rendering |

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
| **Right-click** preview | Opens context menu (theme toggle, zoom controls) |
| **Click** a diagram | Activates zoom mode for that diagram |
| **Scroll wheel** on active diagram | Zoom in/out |
| **Drag** a diagram | Pan the diagram |

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
│   ├── MermaidPreview.cpp   # Plugin main logic
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
│       └── mermaid.min.js   # Embedded mermaid.js (~2MB)
└── bun-renderer/
    ├── package.json
    └── renderer.ts          # Bun-based mermaid renderer
```

## How It Works

1. **Plugin loads** → Registers as an EmEditor plug-in with toolbar button
2. **User clicks button** → Creates a dockable Custom Bar with a child window
3. **WebView2 initializes** → Loads a local HTML shell (`preview.html`) containing CSS and JavaScript
4. **Markdown parsing** → `MarkdownParser` converts editor content to HTML with line-number tracking (`data-line-start`/`data-line-end` attributes)
5. **Mermaid rendering** → Mermaid code blocks become `<div class="mermaid-container">` placeholders; optionally pre-rendered by Bun, otherwise rendered client-side by mermaid.js
6. **Live updates** → `EVENT_MODIFIED` triggers debounced re-render; `EVENT_SCROLL` triggers scroll sync
7. **Bidirectional sync** — Line-number attributes enable precise scroll mapping between editor and preview

## License

MIT

## Credits

- [Mermaid.js](https://mermaid.js.org/) — Diagram rendering engine
- [Microsoft WebView2](https://developer.microsoft.com/en-us/microsoft-edge/webview2/) — Embedded browser control
- [Bun](https://bun.sh/) — JavaScript runtime for server-side rendering
- [EmEditor](https://www.emeditor.com/) — Text editor and plug-in SDK
