# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

MermaidPreview is a native EmEditor plug-in (Windows DLL, C++17 / MSVC) that renders live Markdown + Mermaid diagrams in a dockable Custom Bar. Output is `MermaidPreview.dll`, dropped into EmEditor's `PlugIns` directory.

## Build Commands

The build toolchain is MSVC 2022 Enterprise + Ninja. Use the VS2022 absolute paths from the global instructions, not system PATH:

```bash
# Configure (Release). WebView2 SDK is fetched automatically via FetchContent.
"/c/Program Files/Microsoft Visual Studio/2022/Enterprise/Common7/IDE/CommonExtensions/Microsoft/CMake/CMake/bin/cmake.exe" --preset x64-release

# Build → build/MermaidPreview.dll
"/c/Program Files/Microsoft Visual Studio/2022/Enterprise/Common7/IDE/CommonExtensions/Microsoft/CMake/CMake/bin/cmake.exe" --build build

# Debug (output → build-debug/)
cmake --preset x64-debug && cmake --build build-debug
```

Note: the `CMakePresets.json` already injects the MSVC + Ninja paths into `PATH` for the preset, so plain `cmake.exe` from those preset envs works once invoked through the preset.

There is no automated test harness. The `test/` directory holds Markdown fixtures (e.g. `test/target.md`, `test/docs/`) for manual verification — open them in EmEditor with the DLL loaded.

## Bun Renderer (Optional Component)

Server-side Mermaid pre-rendering lives in `bun-renderer/`:

```bash
cd bun-renderer
bun install                  # mermaid + jsdom
bun renderer.ts              # the long-lived renderer process; talks JSON-lines over stdio
```

`BunRenderer.cpp` spawns this process via `CreateProcessW` and pipes JSON requests on stdin / SVG responses on stdout. If Bun is not installed, the C++ side falls back to client-side rendering inside WebView2 — the plug-in still works.

## Architecture

The DLL is a CETLFrame-based EmEditor plugin (see `include/etlframe.h`, `include/plugin.h` — these are EmEditor SDK headers). Exports are fixed by `exports.def` (`OnCommand`, `QueryStatus`, `OnEvents`, plus the four metadata getters and `PlugInProc`).

```
EmEditor host
  └── PlugInProc / OnCommand / OnEvents          (exports.def)
       └── CMermaidFrame : CETLFrame<CMermaidFrame>   (src/MermaidPreview.{h,cpp})
            ├── Custom Bar lifecycle (OpenCustomBar / CloseCustomBar / Park)
            ├── Auto-open detection (IsMarkdownFile + HasMermaidBlocks)
            ├── Bidirectional scroll sync (anti-feedback flags m_bSyncFromEditor / m_bSyncFromPreview)
            ├── MarkdownParser           ── pure C++, document text → HTML with line-number attrs
            ├── WebView2Manager          ── WebView2 init, JS interop, message channel, parking
            └── BunRenderer (optional)   ── stdio JSON-lines IPC to Bun process
```

Key cross-file flows you must understand before editing:

1. **Render pipeline** (`MermaidPreview::UpdatePreview`):
   `EmEditor doc text` → `MarkdownParser::ConvertToHtml` (each block tagged with `data-line-start` / `data-line-end`) → optional `BunRenderer::RenderBlocks` replaces `<div class="mermaid-container">` placeholders with SVG → `WebView2Manager::RenderContent` ships HTML into the WebView. Mermaid blocks not pre-rendered by Bun fall through to client-side `mermaid.run()` inside the embedded HTML shell built by `WebView2Manager::BuildHtmlPage`.

2. **Parking window optimization** (`m_hwndParking`, `WebView2Manager::Park` / `Reparent`): on `CloseCustomBar`, the WebView2 controller is *not* destroyed — it is reparented into a hidden window owned by `CMermaidFrame`. Reopening calls `Reparent` back into a new host and skips full WebView2 init (~10 ms vs ~800–1500 ms). Any code that destroys the WebView2 must go through `WebView2Manager::Destroy` and respect `m_bDestroyed` to cancel pending async callbacks.

3. **Scroll-sync feedback loop**: `EVENT_SCROLL` from EmEditor and the JS `scroll` callback from WebView2 can each trigger the other. Always set/clear `m_bSyncFromEditor` / `m_bSyncFromPreview` around the dispatch to break the cycle. Same pattern guards `OnPreviewTextEdited` (preview → editor write-back).

4. **Mermaid.js delivery**: `resources/web/mermaid.min.js` (mermaid **11.14.0**, ~3.1 MB UMD bundle) is compiled into the DLL as a resource (`resources/MermaidPreview.rc`, `IDR_MERMAID_JS RCDATA`) and extracted to disk on first use by `WebView2Manager::ExtractMermaidJs`. `BuildHtmlPage` references the extracted file via `file://` so WebView2 can cache it. All four `mermaid.initialize` sites in `BuildHtmlPage` go through one JS helper `_mmdInit(theme, look)` — when adding new mermaid config flags (e.g. `look:'neo'`, `architecture.randomize`, timeline directional control), edit that helper only. The Bun renderer (`bun-renderer/renderer.ts`) accepts `theme` and `look` per render request and applies the same config schema.

5. **Security boundaries** to preserve when editing: `MarkdownParser::IsSafeUrl` (blocks `javascript:` / `data:` / `vbscript:` schemes — must be called for every link / image URL), WebView2 host-object injection is disabled in `WebView2Manager`, and JS-bound integers (line numbers from preview) are validated before passing back to EmEditor APIs.

## Source Files (size guide)

The C++ source is concentrated in four files; line counts are useful for choosing read strategies:

| File | LOC | Role |
|------|-----|------|
| `src/MermaidPreview.cpp` | ~978 | Plugin frame, Custom Bar, event dispatch, settings |
| `src/MarkdownParser.cpp` | ~962 | Markdown→HTML, mermaid block extraction, table parser |
| `src/WebView2Manager.cpp` | ~940 | WebView2 lifecycle, JS message channel, parking, HTML shell |
| `src/BunRenderer.cpp` | ~442 | Bun process management, JSON-lines IPC |
| `src/DllMain.cpp` | ~32 | DLL entry only |

When asked to read these in full, prefer targeted Grep / line-range Reads over a 2000-line dump.

## Conventions

- C++17, `/EHsc /W3`, static CRT (`MultiThreaded[Debug]`), `UNICODE` / `_UNICODE` / `NOMINMAX`.
- All EmEditor APIs and Win32 strings are wide (`std::wstring`); UTF-8 conversion happens only at the Bun stdio boundary (`BunRenderer::WtoU8` / `U8toW`).
- COM lifetime: `Microsoft::WRL::ComPtr` everywhere. Event tokens (`m_navCompletedToken`, `m_webMessageToken`) must be unsubscribed in `WebView2Manager::Destroy`.
- Settings live in the registry; `LoadSettings` / `SaveSettings` in `CMermaidFrame` is the single chokepoint.
- `_ALLOW_MULTIPLE_INSTANCES = FALSE` and `_USE_CUSTOM_BAR = TRUE` in `CMermaidFrame` — do not flip these without re-checking parking-window assumptions.
