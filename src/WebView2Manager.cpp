#include "WebView2Manager.h"
#include "resource.h"
#include <shlobj.h>
#include <sstream>
#include <climits>

using namespace Microsoft::WRL;

// Get the DLL's HMODULE (for loading resources)
extern HINSTANCE EEGetInstanceHandle();

// HTML cache version tag — increment when BuildHtmlPage() content changes
static const char* kHtmlVersionTag = "<!-- MermaidPreview-v14 -->";

WebView2Manager::WebView2Manager() = default;

WebView2Manager::~WebView2Manager()
{
    Destroy();
}

// ============================================================================
// GetResourceDir - %LOCALAPPDATA%\MermaidPreview\web
// ============================================================================
std::wstring WebView2Manager::GetResourceDir() const
{
    WCHAR appData[MAX_PATH] = {};
    if (SUCCEEDED(SHGetFolderPathW(nullptr, CSIDL_LOCAL_APPDATA, nullptr, 0, appData))) {
        std::wstring dir = appData;
        dir += L"\\MermaidPreview\\web";
        return dir;
    }
    return L"";
}

// ============================================================================
// ExtractMermaidJs - Extract mermaid.min.js from DLL resource to disk
// ============================================================================
bool WebView2Manager::ExtractMermaidJs()
{
    m_resourceDir = GetResourceDir();
    if (m_resourceDir.empty())
        return false;

    // Create directory (recursively)
    std::wstring parentDir = m_resourceDir.substr(0, m_resourceDir.rfind(L'\\'));
    CreateDirectoryW(parentDir.c_str(), nullptr);
    CreateDirectoryW(m_resourceDir.c_str(), nullptr);

    std::wstring mermaidPath = m_resourceDir + L"\\mermaid.min.js";

    // Check if already extracted (skip if file exists and has reasonable size)
    WIN32_FILE_ATTRIBUTE_DATA fileInfo = {};
    if (GetFileAttributesExW(mermaidPath.c_str(), GetFileExInfoStandard, &fileInfo)) {
        if (fileInfo.nFileSizeLow > 100000) // > 100KB means valid
            return true;
    }

    // Load from DLL resource
    HINSTANCE hInst = EEGetInstanceHandle();
    HRSRC hRes = FindResourceW(hInst, MAKEINTRESOURCEW(IDR_MERMAID_JS), RT_RCDATA);
    if (!hRes)
        return false;

    HGLOBAL hData = LoadResource(hInst, hRes);
    if (!hData)
        return false;

    DWORD size = SizeofResource(hInst, hRes);
    const void* data = LockResource(hData);
    if (!data || size == 0)
        return false;

    // Write to file. If WriteFile fails or only writes part of the buffer
    // (disk full, antivirus interference, ...), delete the half-written
    // file so the next launch re-extracts a fresh copy instead of loading
    // a truncated mermaid.min.js (which would silently corrupt rendering).
    HANDLE hFile = CreateFileW(mermaidPath.c_str(), GENERIC_WRITE, 0, nullptr,
                               CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (hFile == INVALID_HANDLE_VALUE)
        return false;

    DWORD written = 0;
    BOOL writeOk = WriteFile(hFile, data, size, &written, nullptr);
    CloseHandle(hFile);

    if (!writeOk || written != size) {
        DeleteFileW(mermaidPath.c_str());
        return false;
    }
    return true;
}

// ============================================================================
// BuildHtmlPage - Minimal HTML shell (no CDN, mermaid.js loaded locally)
// ============================================================================
std::wstring WebView2Manager::BuildHtmlPage() const
{
    std::wstring mermaidSrc = L"mermaid.min.js";

    // Split into multiple parts to stay under MSVC 16KB string literal limit
    // and avoid raw-string delimiter conflicts with JS code.

    // --- Part 1: CSS ---
    std::wstring html = LR"P1(<!-- MermaidPreview-v14 -->
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<style>
  * { box-sizing: border-box; }
  html { font-size: 14px; }
  body {
    margin: 0; padding: 16px 20px;
    font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif;
    line-height: 1.6; overflow-y: auto; overflow-x: hidden;
    word-wrap: break-word; transition: background-color 0.15s, color 0.15s;
  }
  body.light { background: #fff; color: #24292f; }
  body.light a { color: #0969da; }
  body.light hr { border-color: #d0d7de; }
  body.light blockquote { border-left-color: #d0d7de; color: #57606a; }
  body.light code { background: #eff1f3; }
  body.light pre { background: #f6f8fa; border-color: #d0d7de; }
  body.light table th { background: #f6f8fa; }
  body.light table td, body.light table th { border-color: #d0d7de; }
  body.light .mermaid-error { color: #cf222e; background: #ffebe9; }
  body.dark { background: #1e1e1e; color: #d4d4d4; }
  body.dark a { color: #58a6ff; }
  body.dark hr { border-color: #444; }
  body.dark blockquote { border-left-color: #444; color: #8b949e; }
  body.dark code { background: #2d2d2d; }
  body.dark pre { background: #252526; border-color: #444; }
  body.dark table th { background: #2d2d2d; }
  body.dark table td, body.dark table th { border-color: #444; }
  body.dark .mermaid-error { color: #f85149; background: #3e2723; }
  h1, h2, h3, h4, h5, h6 { margin: 1em 0 0.5em; font-weight: 600; line-height: 1.25; }
  h1 { font-size: 1.75em; padding-bottom: 0.3em; border-bottom: 1px solid; border-color: inherit; }
  h2 { font-size: 1.45em; padding-bottom: 0.25em; border-bottom: 1px solid; border-color: inherit; }
  h3 { font-size: 1.2em; } h4 { font-size: 1em; }
  p { margin: 0.6em 0; }
  a { text-decoration: none; } a:hover { text-decoration: underline; }
  hr { border: 0; border-top: 1px solid; margin: 1.2em 0; }
  img { max-width: 100%; height: auto; }
  ul, ol { margin: 0.5em 0; padding-left: 2em; }
  li { margin: 0.25em 0; } li > ul, li > ol { margin: 0.15em 0; }
  blockquote { margin: 0.5em 0; padding: 0.5em 1em; border-left: 4px solid; font-style: italic; }
  blockquote p { margin: 0.3em 0; }
  code { font-family: 'Cascadia Code','Fira Code',Consolas,'Courier New',monospace; font-size: 0.88em; padding: 0.15em 0.35em; border-radius: 3px; }
  pre { margin: 0.8em 0; padding: 12px 16px; border: 1px solid; border-radius: 6px; overflow-x: auto; line-height: 1.45; }
  pre code { padding: 0; background: none; font-size: 0.88em; }
  table { border-collapse: collapse; width: 100%; margin: 0.8em 0; font-size: 0.92em; }
  th, td { padding: 6px 12px; border: 1px solid; text-align: left; } th { font-weight: 600; }
  .task-list-item { list-style: none; margin-left: -1.5em; }
  .task-list-item input[type="checkbox"] { margin-right: 0.5em; }
  /* Default = horizontal scroll when SVG is naturally wider than the pane.
     Pan/zoom modes switch back to overflow:hidden so the CSS transform
     doesn't push surrounding content. (Audit follow-up: wide RE flowcharts
     used to be silently clipped on the right.) */
  .mermaid-container { margin: 1em 0; overflow-x: auto; overflow-y: hidden; cursor: grab; text-align: center; position: relative; border-radius: 4px; min-height: 40px; }
  .mermaid-container.dragging,
  .mermaid-container.zoom-active { overflow: hidden; }
  .mermaid-container.dragging { cursor: grabbing; }
  .mermaid-container.zoom-active { outline: 2px dashed #0969da; outline-offset: 2px; cursor: zoom-in; }
  body.dark .mermaid-container.zoom-active { outline-color: #58a6ff; }
  .mermaid-container.zoom-active .svg-zoom-badge { opacity:1 !important; background:rgba(9,105,218,0.7); }
  body.dark .mermaid-container.zoom-active .svg-zoom-badge { background:rgba(88,166,255,0.7); }
  .mermaid-container svg { display: inline-block; max-width: 100%; height: auto; transform-origin: 0 0; user-select: none; }
  /* When the user is actively zooming/panning, let the SVG render at its
     natural size so the in-pane scroll bar is the right escape hatch. */
  .mermaid-container.zoom-active svg { max-width: none; }
  /* Edge-label visibility hardening (mermaid 11.14 emits 0.5–0.8 alpha
     and paints labels BEFORE nodes — see liftEdgeLabels for the painter
     fix). Force fully-opaque background + a stroke halo on the text so
     even partial overlaps stay readable. */
  .mermaid-container svg .edgeLabel,
  .mermaid-container svg .labelBkg { background-color: #ffffff !important; }
  .mermaid-container svg .edgeLabel rect,
  .mermaid-container svg .edgeLabel foreignObject > div { background-color: #ffffff !important; opacity: 1 !important; }
  .mermaid-container svg .edgeLabel text,
  .mermaid-container svg .edgeLabel span { paint-order: stroke; stroke: #ffffff; stroke-width: 3px; }
  body.dark .mermaid-container svg .edgeLabel,
  body.dark .mermaid-container svg .labelBkg { background-color: #1e1e1e !important; }
  body.dark .mermaid-container svg .edgeLabel rect,
  body.dark .mermaid-container svg .edgeLabel foreignObject > div { background-color: #1e1e1e !important; opacity: 1 !important; }
  body.dark .mermaid-container svg .edgeLabel text,
  body.dark .mermaid-container svg .edgeLabel span { paint-order: stroke; stroke: #1e1e1e; stroke-width: 3px; }
  .mermaid-error { padding: 8px 12px; border-radius: 4px; font-size: 0.85em; margin: 0.5em 0; }
  #ctx-menu { display:none; position:fixed; z-index:9999; min-width:180px; padding:4px 0; background:#fff; border:1px solid #d0d7de; border-radius:6px; box-shadow:0 4px 16px rgba(0,0,0,0.12); font-size:13px; }
  body.dark #ctx-menu { background:#2d2d2d; border-color:#444; }
  .ctx-item { padding:6px 20px; cursor:pointer; white-space:nowrap; }
  .ctx-item:hover { background:#0969da; color:#fff; }
  body.dark .ctx-item:hover { background:#58a6ff; color:#1e1e1e; }
  .ctx-item.ctx-active { font-weight:600; }
  .ctx-item.ctx-active::before { content:'> '; font-weight:400; }
  .ctx-sep { margin:4px 12px; border:0; border-top:1px solid #d0d7de; }
  body.dark .ctx-sep { border-color:#444; }
  .svg-zoom-badge { position:absolute; top:4px; right:4px; background:rgba(0,0,0,0.5); color:#fff; font-size:11px; padding:1px 6px; border-radius:3px; pointer-events:none; opacity:0; transition:opacity 0.2s; }
  .mermaid-container:hover .svg-zoom-badge { opacity:1; }
  /* Expand-to-pane button (sits next to .svg-zoom-badge) */
  .mermaid-expand-btn { position:absolute; top:4px; right:50px; background:rgba(0,0,0,0.5); color:#fff; font-size:13px; line-height:1; padding:3px 7px; border:0; border-radius:3px; cursor:pointer; opacity:0; transition:opacity 0.2s; font-family:inherit; }
  .mermaid-container:hover .mermaid-expand-btn,
  .mermaid-container.mp-expanded .mermaid-expand-btn { opacity:1; }
  .mermaid-expand-btn:hover { background:rgba(9,105,218,0.85); }
  body.dark .mermaid-expand-btn:hover { background:rgba(88,166,255,0.85); }
  /* Fullscreen takeover: one container fills the entire preview pane,
     scroll bars handle anything bigger. body.mp-fs hides the rest. */
  body.mp-fs > #content > :not(.mermaid-container.mp-expanded) { display:none; }
  .mermaid-container.mp-expanded { position:fixed; inset:0; z-index:10000; margin:0; border-radius:0; overflow:auto; background:#ffffff; padding:24px; }
  body.dark .mermaid-container.mp-expanded { background:#1e1e1e; }
  .mermaid-container.mp-expanded svg { max-width:none; }
  .loading,.empty { text-align:center; color:#999; padding:40px 16px; font-size:14px; }
  body.dark .loading, body.dark .empty { color:#666; }
  #content > :first-child { margin-top: 0; }
  [data-line-start] { cursor:default; border-radius:3px; transition:outline 0.15s; }
  [data-line-start]:hover { outline:1px dashed #ccc; }
  body.dark [data-line-start]:hover { outline-color:#555; }
  .editing { outline:2px solid #0969da !important; background:rgba(9,105,218,0.05); padding:2px 4px; }
  body.dark .editing { outline-color:#58a6ff !important; background:rgba(88,166,255,0.05); }
)P1";

    // --- Part 1a: HTML body ---
    html += LR"P1a(
</style>
</head>
<body class="light">
  <div id="content"><div class="loading">Loading...</div></div>
  <div id="ctx-menu">
    <div class="ctx-item ctx-theme" data-dark="false">Light Theme</div>
    <div class="ctx-item ctx-theme" data-dark="true">Dark Theme</div>
    <div class="ctx-sep"></div>
    <div class="ctx-item ctx-action" data-action="zoom-in">Zoom In</div>
    <div class="ctx-item ctx-action" data-action="zoom-out">Zoom Out</div>
    <div class="ctx-item ctx-action" data-action="reset">Reset View</div>
    <div class="ctx-item ctx-action" data-action="fit-width">Fit to Width</div>
    <div class="ctx-item ctx-action" data-action="actual-size">Actual Size</div>
    <div class="ctx-sep"></div>
    <div class="ctx-item ctx-action" data-action="font-up">Font Size +</div>
    <div class="ctx-item ctx-action" data-action="font-down">Font Size -</div>
    <div class="ctx-item ctx-action" data-action="font-reset">Font Size Reset</div>
  </div>
)P1a";

    // --- Part 1b: Script opening + mermaid loader ---
    html += LR"P1b(  <script>
    var mermaidReady = false;
    var renderedMermaidSrcs = {};
    var zoomTarget = null;
    var _pendingRender = null;

    // Single chokepoint for mermaid.initialize (11.14+ feature surface).
    // theme: 'default' | 'dark' | 'neutral' | 'forest' | 'base'
    // look : 'classic' (default) | 'neo' | 'handDrawn'
    //
    // The flowchart spacing values are intentionally generous (default 50/50):
    // RE/architecture diagrams pack many subgraph layers and long edge labels,
    // and mermaid 11.14's dagre layout otherwise jams labels into nodes.
    // 'basis' curves give labels a smoother corridor than the default linear.
    function _mmdInit(theme, look) {
      mermaid.initialize({
        startOnLoad: false,
        securityLevel: 'strict',
        theme: theme || 'default',
        look: look || 'classic',
        flowchart: { useMaxWidth: true, nodeSpacing: 60, rankSpacing: 80, curve: 'basis' },
        sequence: { useMaxWidth: true },
        architecture: { randomize: false }
      });
    }

    // Lift <g.edgeLabels> after <g.nodes> so labels paint on top — mermaid's
    // default DOM order paints nodes over labels, which is what produced the
    // overlap the user reported in the RE flowchart.
    function _liftEdgeLabels(root) {
      if (!root) return;
      var svgs = root.querySelectorAll ? root.querySelectorAll('svg') : [];
      for (var k = 0; k < svgs.length; k++) {
        var rootG = svgs[k].querySelector(':scope > g');
        if (!rootG) continue;
        var labelsG = rootG.querySelector(':scope > g.edgeLabels');
        if (labelsG) rootG.appendChild(labelsG); // last child = drawn last = top-most
      }
    }

    // Inject (or refresh) the ⛶ expand button on a mermaid container.
    function _addExpandBtn(container) {
      if (!container || container.querySelector(':scope > .mermaid-expand-btn')) return;
      var btn = document.createElement('button');
      btn.className = 'mermaid-expand-btn';
      btn.type = 'button';
      btn.title = 'Expand to full preview (Esc to exit)';
      btn.textContent = '⛶';
      btn.addEventListener('mousedown', function(e){ e.stopPropagation(); });
      btn.addEventListener('click', function(e){
        e.stopPropagation(); e.preventDefault();
        toggleMermaidExpand(container);
      });
      container.appendChild(btn);
    }

    function toggleMermaidExpand(container) {
      var expanded = container.classList.toggle('mp-expanded');
      // Toggle body class only when *some* container is expanded.
      if (expanded) {
        // Make sure no other container is left expanded.
        document.querySelectorAll('.mermaid-container.mp-expanded').forEach(function(c){
          if (c !== container) c.classList.remove('mp-expanded');
        });
        document.body.classList.add('mp-fs');
      } else {
        document.body.classList.remove('mp-fs');
      }
    }

    // Global Esc → exit fullscreen if any container is expanded.
    document.addEventListener('keydown', function(e){
      if (e.key === 'Escape' && document.body.classList.contains('mp-fs')) {
        document.querySelectorAll('.mermaid-container.mp-expanded').forEach(function(c){
          c.classList.remove('mp-expanded');
        });
        document.body.classList.remove('mp-fs');
      }
    });

    // Async load mermaid.min.js — does not block NavigationCompleted
    (function(){
      var s = document.createElement('script');
      s.src = ')P1b" + mermaidSrc + LR"P2(';
      s.onload = function() {
        _mmdInit('default');
        mermaidReady = true;
        if (_pendingRender) { _processPendingMermaid(); }
      };
      document.head.appendChild(s);
    })();

    async function _processPendingMermaid() {
      if (!_pendingRender) return;
      var isDark = (_pendingRender.theme === 'dark');
      _mmdInit(isDark ? 'dark' : 'default');
      _pendingRender = null;
      var container = document.getElementById('content');
      var placeholders = container.querySelectorAll('.mermaid-container[data-mermaid-src]');
      var newSrcs = {};
      for (var i = 0; i < placeholders.length; i++) {
        var el = placeholders[i];
        var src = decodeURIComponent(el.getAttribute('data-mermaid-src'));
        var id = el.getAttribute('data-mermaid-id') || ('mmd-'+i+'-'+Date.now());
        try {
          var result = await mermaid.render(id, src);
          el.innerHTML = result.svg;
          newSrcs[id] = { src: src, svg: result.svg };
        } catch(e) {
          el.innerHTML = '<div class="mermaid-error">Mermaid: '+String(e.message||e).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;')+'</div>';
          newSrcs[id] = { src: src, svg: el.innerHTML };
        }
      }
      renderedMermaidSrcs = newSrcs;
      container.querySelectorAll('.mermaid-container').forEach(function(c){ initSvgDrag(c); _addExpandBtn(c); });
      _liftEdgeLabels(container);
    }

    window.renderContent = async function(htmlContent, theme) {
      var isDark = (theme === 'dark');
      document.body.className = isDark ? 'dark' : 'light';
      if (mermaidReady) {
        _mmdInit(isDark ? 'dark' : 'default');
      }
      var container = document.getElementById('content');
      container.innerHTML = htmlContent;
      if (mermaidReady) {
        var placeholders = container.querySelectorAll('.mermaid-container[data-mermaid-src]');
        var newSrcs = {};
        for (var i = 0; i < placeholders.length; i++) {
          var el = placeholders[i];
          var src = decodeURIComponent(el.getAttribute('data-mermaid-src'));
          var id = el.getAttribute('data-mermaid-id') || ('mmd-'+i+'-'+Date.now());
          if (renderedMermaidSrcs[id] && renderedMermaidSrcs[id].src === src) {
            el.innerHTML = renderedMermaidSrcs[id].svg; newSrcs[id] = renderedMermaidSrcs[id]; continue;
          }
          try {
            var result = await mermaid.render(id, src);
            el.innerHTML = result.svg;
            newSrcs[id] = { src: src, svg: result.svg };
          } catch(e) {
            el.innerHTML = '<div class="mermaid-error">Mermaid: '+String(e.message||e).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;')+'</div>';
            newSrcs[id] = { src: src, svg: el.innerHTML };
          }
        }
        renderedMermaidSrcs = newSrcs;
      } else {
        _pendingRender = { theme: theme };
      }
      container.querySelectorAll('.mermaid-container').forEach(function(c){ initSvgDrag(c); _addExpandBtn(c); });
      _liftEdgeLabels(container);
    };
    window.setTheme = function(dark) { document.body.className = dark ? 'dark' : 'light'; };
    window.clearContent = function() { document.getElementById('content').innerHTML = '<div class="empty">Open a Markdown file to preview</div>'; renderedMermaidSrcs = {}; _pendingRender = null; };

    // ===== Font Size =====
    var _fontSize = 14;
    window.setFontSize = function(n) {
      n = Math.max(8, Math.min(32, n));
      _fontSize = n;
      document.documentElement.style.fontSize = n + 'px';
    };
    function fontSizeChange(delta) {
      var n = Math.max(8, Math.min(32, _fontSize + delta));
      if (n === _fontSize) return;
      _fontSize = n;
      document.documentElement.style.fontSize = n + 'px';
      if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({type:'fontSize', size:n});
      }
    }
    function fontSizeReset() {
      _fontSize = 14;
      document.documentElement.style.fontSize = '14px';
      if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({type:'fontSize', size:14});
      }
    }
)P2";

    // --- Part 2: Scroll Sync + Editing + Context Menu + SVG Drag JS ---
    html += LR"P3(
    // ===== Scroll Sync =====
    var _syncFromCpp = false;
    var _scrollTimer = 0;
    window.scrollToLine = function(line) {
      _syncFromCpp = true;
      var els = document.querySelectorAll('[data-line-start]');
      var best = null;
      for (var i = 0; i < els.length; i++) {
        var ls = parseInt(els[i].getAttribute('data-line-start'));
        if (ls <= line) best = els[i];
        if (ls >= line) { best = els[i]; break; }
      }
      if (best) best.scrollIntoView({ behavior:'auto', block:'start' });
      setTimeout(function(){ _syncFromCpp = false; }, 150);
    };
    window.addEventListener('scroll', function() {
      if (_syncFromCpp) return;
      clearTimeout(_scrollTimer);
      _scrollTimer = setTimeout(function() {
        var els = document.querySelectorAll('[data-line-start]');
        var topLine = 0;
        for (var i = 0; i < els.length; i++) {
          var rect = els[i].getBoundingClientRect();
          if (rect.top >= 0) { topLine = parseInt(els[i].getAttribute('data-line-start')); break; }
        }
        if (window.chrome && window.chrome.webview) {
          window.chrome.webview.postMessage({type:'syncScroll', line: topLine});
        }
      }, 50);
    });

    // ===== Editing =====
    // Double-click → contentEditable. Two paths:
    //   (a) markdown elements with data-line-start → POST `edit` with the
    //       full line range so OnPreviewTextEdited can replace it.
    //   (b) mermaid SVG node labels (M2) → POST `editMermaidNode` with
    //       blockId+nodeId+newLabel; C++ resolves which line in the block
    //       carries the matching label and rewrites just that range.
    document.getElementById('content').addEventListener('dblclick', function(e) {
      // --- (b) Mermaid node label path ---
      var mc = e.target.closest && e.target.closest('.mermaid-container');
      if (mc) {
        var nodeG = e.target.closest('g.node');
        if (!nodeG) return;
        var blockId = mc.getAttribute('data-mermaid-id');
        // Mermaid 11.14 emits id="<svgId>-flowchart-<nodeId>-<n>". Strip
        // the "<svgId>-flowchart-" prefix and trailing "-<n>" to recover
        // the original mermaid identifier.
        var rawId = nodeG.getAttribute('id') || '';
        var m = rawId.match(/-flowchart-(.+)-\d+$/);
        if (!m || !blockId) return;
        var nodeId = m[1];
        // Find the human-editable target inside the foreignObject.
        var p = nodeG.querySelector('foreignObject p, foreignObject span.nodeLabel, foreignObject div');
        if (!p || p.classList.contains('editing')) return;
        // Convert any rendered <br> back to a real newline so the user
        // sees one line per source line and can edit naturally.
        if (p.querySelector && p.querySelector('br')) {
          p.innerHTML = p.innerHTML.replace(/<br\s*\/?>/gi, '\n');
        }
        p.contentEditable = 'true';
        p.classList.add('editing');
        p.style.whiteSpace = 'pre-wrap';
        p._origText = p.textContent;
        // Stop the click bubbling up to the navigateToLine handler that
        // also lives on .mermaid-container.
        e.stopPropagation();
        e.preventDefault();
        // Defer focus + select-all so the browser doesn't immediately
        // clear the selection from the dblclick.
        setTimeout(function(){
          p.focus();
          try { var rng = document.createRange(); rng.selectNodeContents(p);
                var sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(rng); } catch(_){}
        }, 0);
        function mmdSave() {
          var txt = (p.textContent || '');
          if (txt !== p._origText) {
            window.chrome.webview.postMessage({
              type:'editMermaidNode',
              blockId: blockId,
              nodeId: nodeId,
              newLabel: txt
            });
          }
          mmdFinish();
        }
        function mmdFinish() {
          p.contentEditable = 'false';
          p.classList.remove('editing');
          p.style.whiteSpace = '';
          p.removeEventListener('blur', mmdBlur);
          p.removeEventListener('keydown', mmdKey);
        }
        function mmdBlur() { mmdSave(); }
        function mmdKey(ev) {
          if (ev.key === 'Escape') { p.textContent = p._origText; mmdFinish(); }
          else if (ev.key === 'Enter' && !ev.shiftKey) { ev.preventDefault(); mmdSave(); }
        }
        p.addEventListener('blur', mmdBlur);
        p.addEventListener('keydown', mmdKey);
        return;
      }

      // --- (a) Plain markdown path (existing behaviour) ---
      var el = e.target;
      while (el && el !== this) { if (el.hasAttribute('data-line-start')) break; el = el.parentElement; }
      if (!el || !el.hasAttribute('data-line-start')) return;
      if (el.classList.contains('editing')) return;
      el.contentEditable = 'true'; el.classList.add('editing'); el.focus();
      el._origText = el.textContent;
      function save() {
        var txt = el.textContent, ls = parseInt(el.getAttribute('data-line-start')), le = parseInt(el.getAttribute('data-line-end'));
        if (txt !== el._origText) { window.chrome.webview.postMessage({type:'edit',lineStart:ls,lineEnd:le,newText:txt}); }
        finish();
      }
      function finish() { el.contentEditable='false'; el.classList.remove('editing'); el.removeEventListener('blur',onBlur); el.removeEventListener('keydown',onKey); }
      function onBlur() { save(); }
      function onKey(ev) { if (ev.key==='Escape'){el.textContent=el._origText;finish();} else if(ev.key==='Enter'&&!ev.shiftKey){ev.preventDefault();save();} }
      el.addEventListener('blur', onBlur); el.addEventListener('keydown', onKey);
    });
)P3";

    // --- Part 3b: Context Menu + SVG Drag JS ---
    html += LR"P4(
    // ===== Context Menu =====
    var ctxMenu = document.getElementById('ctx-menu');
    document.addEventListener('contextmenu', function(e) {
      e.preventDefault();
      var isDark = document.body.classList.contains('dark');
      document.querySelectorAll('.ctx-theme').forEach(function(t){ t.classList.toggle('ctx-active',(t.dataset.dark==='true')===isDark); });
      ctxMenu.style.display = 'block';
      var x=e.clientX, y=e.clientY, r=ctxMenu.getBoundingClientRect();
      if(x+r.width>window.innerWidth) x=window.innerWidth-r.width-4;
      if(y+r.height>window.innerHeight) y=window.innerHeight-r.height-4;
      ctxMenu.style.left=x+'px'; ctxMenu.style.top=y+'px';
    });
    document.addEventListener('click', function(e){
      ctxMenu.style.display='none';
      if(zoomTarget && !e.target.closest('.mermaid-container')){
        zoomTarget.classList.remove('zoom-active'); zoomTarget=null;
      }
      // Handle internal anchor links (#section) and relative file links
      var anchor = e.target.closest('a[href]');
      if(anchor){
        var href = anchor.getAttribute('href');
        if(href && href.charAt(0)==='#'){
          e.preventDefault();
          var target = document.getElementById(href.substring(1));
          if(target) target.scrollIntoView({behavior:'instant',block:'start'});
        } else if(href && !href.match(/^[a-zA-Z][a-zA-Z0-9+.\-]*:/) && window.chrome && window.chrome.webview){
          e.preventDefault();
          window.chrome.webview.postMessage({type:'openFile', path:href});
        }
      }
    });
    window.addEventListener('blur', function(){ ctxMenu.style.display='none'; });
    document.querySelectorAll('.ctx-theme').forEach(function(el){ el.addEventListener('click',function(){ switchTheme(this.dataset.dark==='true'); ctxMenu.style.display='none'; }); });
    document.querySelectorAll('.ctx-action').forEach(function(el){ el.addEventListener('click',function(){
      var a=this.dataset.action;
      if(a==='zoom-in') svgZoomAll(1.25);
      else if(a==='zoom-out') svgZoomAll(0.8);
      else if(a==='reset') svgResetAll();
      else if(a==='fit-width') svgFitWidthAll();
      else if(a==='actual-size') svgActualSizeAll();
      else if(a==='font-up') fontSizeChange(2);
      else if(a==='font-down') fontSizeChange(-2);
      else if(a==='font-reset') fontSizeReset();
      ctxMenu.style.display='none';
    }); });

    // Fit to Width: clear any pan/zoom transform and re-clamp SVG to
    // container width via the default `max-width: 100%` rule (i.e. drop
    // the .zoom-active state, which is what `max-width: none` came from).
    function svgFitWidthAll(){
      document.querySelectorAll('.mermaid-container').forEach(function(c){
        c.classList.remove('zoom-active');
        var svg = c.querySelector('svg');
        if(svg){ applySvgState(svg, {tx:0, ty:0, scale:1}); }
      });
      if(zoomTarget){ zoomTarget=null; }
    }
    // Actual Size: render at SVG's natural width and let the container's
    // default overflow-x:auto scroll. We drop .zoom-active so the
    // overflow:hidden side-effect doesn't kick in, and set maxWidth:none
    // inline (cleared on next render, which is the desired session-scope).
    function svgActualSizeAll(){
      document.querySelectorAll('.mermaid-container').forEach(function(c){
        c.classList.remove('zoom-active');
        var svg = c.querySelector('svg');
        if(svg){ svg.style.maxWidth = 'none'; applySvgState(svg, {tx:0, ty:0, scale:1}); }
      });
      if(zoomTarget){ zoomTarget=null; }
    }

    function switchTheme(dark) {
      document.body.className = dark ? 'dark' : 'light';
      if (mermaidReady) {
        _mmdInit(dark ? 'dark' : 'default');
        document.querySelectorAll('.mermaid-container[data-mermaid-src]').forEach(function(el,idx){
          var src = decodeURIComponent(el.getAttribute('data-mermaid-src'));
          mermaid.render('ts-'+idx+'-'+Date.now(), src).then(function(r){ el.innerHTML=r.svg; initSvgDrag(el); _addExpandBtn(el); _liftEdgeLabels(el); }).catch(function(){});
        });
      }
      window.chrome.webview.postMessage({type:'theme',dark:dark});
    }

    // ===== SVG Drag / Pan / Zoom =====
    function getSvgState(svg) { return { tx:parseFloat(svg.dataset.tx||0), ty:parseFloat(svg.dataset.ty||0), scale:parseFloat(svg.dataset.scale||1) }; }
    function applySvgState(svg, s) {
      svg.dataset.tx=s.tx; svg.dataset.ty=s.ty; svg.dataset.scale=s.scale;
      svg.style.transform='translate('+s.tx+'px,'+s.ty+'px) scale('+s.scale+')';
      var badge=svg.parentElement?svg.parentElement.querySelector('.svg-zoom-badge'):null;
      if(badge){ badge.textContent=Math.round(s.scale*100)+'%'; badge.style.opacity=(s.scale===1&&s.tx===0&&s.ty===0)?'':'1'; }
    }
    function initSvgDrag(container) {
      var svg=container.querySelector('svg'); if(!svg) return;
      svg.dataset.tx='0'; svg.dataset.ty='0'; svg.dataset.scale='1'; svg.style.transformOrigin='0 0';
      if(!container.querySelector('.svg-zoom-badge')){ var b=document.createElement('span'); b.className='svg-zoom-badge'; b.textContent='100%'; container.appendChild(b); }
    }
    document.getElementById('content').addEventListener('mousedown', function(e) {
      if(e.button!==0) return;
      var container=e.target.closest('.mermaid-container'); if(!container) return;
      var svg=container.querySelector('svg'); if(!svg) return;
      e.preventDefault();
      var startX=e.clientX, startY=e.clientY, moved=false, wasCtrl=e.ctrlKey;
      var s=getSvgState(svg), baseTx=s.tx, baseTy=s.ty;
      function onMove(ev){
        var dx=ev.clientX-startX, dy=ev.clientY-startY;
        if(Math.abs(dx)>3||Math.abs(dy)>3) moved=true;
        if(moved){ container.classList.add('dragging'); s.tx=baseTx+dx; s.ty=baseTy+dy; applySvgState(svg,s); }
      }
      function onUp(){
        container.classList.remove('dragging');
        document.removeEventListener('mousemove',onMove); document.removeEventListener('mouseup',onUp);
        if(!moved){
          if(wasCtrl){
            // Ctrl+Click: toggle zoom-active mode
            if(zoomTarget===container){ zoomTarget=null; container.classList.remove('zoom-active'); }
            else { if(zoomTarget) zoomTarget.classList.remove('zoom-active'); zoomTarget=container; container.classList.add('zoom-active'); }
          } else {
            // Regular click: sync editor to this mermaid block's source line
            var ls = container.getAttribute('data-line-start');
            if(ls && window.chrome && window.chrome.webview){
              window.chrome.webview.postMessage({type:'navigateToLine', line:parseInt(ls)});
            }
          }
        }
      }
      document.addEventListener('mousemove',onMove); document.addEventListener('mouseup',onUp);
    });
    document.getElementById('content').addEventListener('wheel', function(e) {
      var container=e.target.closest('.mermaid-container');
      if(!container || (!e.ctrlKey && container!==zoomTarget)) return;
      var svg=container.querySelector('svg'); if(!svg) return;
      e.preventDefault(); var s=getSvgState(svg), factor=e.deltaY>0?0.9:1.1;
      var ns=Math.max(0.2,Math.min(5,s.scale*factor));
      var rect=container.getBoundingClientRect(), mx=e.clientX-rect.left, my=e.clientY-rect.top;
      s.tx=mx-(mx-s.tx)*(ns/s.scale); s.ty=my-(my-s.ty)*(ns/s.scale); s.scale=ns;
      applySvgState(svg,s);
    }, {passive:false});
    function svgZoomAll(f){ document.querySelectorAll('.mermaid-container svg').forEach(function(svg){ var s=getSvgState(svg); s.scale=Math.max(0.2,Math.min(5,s.scale*f)); applySvgState(svg,s); }); }
    function svgResetAll(){ document.querySelectorAll('.mermaid-container svg').forEach(function(svg){ applySvgState(svg,{tx:0,ty:0,scale:1}); }); }
  </script>
</body>
</html>)P4";

    return html;
}

// ============================================================================
// EscapeForJS
// ============================================================================
std::wstring WebView2Manager::EscapeForJS(const std::wstring& input)
{
    std::wstring result;
    result.reserve(input.size() + input.size() / 4);
    for (wchar_t ch : input) {
        switch (ch) {
        case L'\0': result += L"\\u0000"; break;         // Avoid silent JS string truncation
        case L'\\': result += L"\\\\"; break;
        case L'\'': result += L"\\'";  break;
        case L'"':  result += L"\\\""; break;
        case L'\n': result += L"\\n";  break;
        case L'\r': result += L"\\r";  break;
        case L'\t': result += L"\\t";  break;
        case L'<':  result += L"\\x3C"; break;
        case L'>':  result += L"\\x3E"; break;           // Prevent </script> injection
        case L'\u2028': result += L"\\u2028"; break;     // Line Separator
        case L'\u2029': result += L"\\u2029"; break;     // Paragraph Separator
        default:    result += ch;       break;
        }
    }
    return result;
}

// ============================================================================
// Initialize
// ============================================================================
HRESULT WebView2Manager::Initialize(HWND hwndParent, std::function<void()> onReady)
{
    m_hwndParent = hwndParent;
    m_onReady = std::move(onReady);
    m_bDestroyed = false;

    // Extract mermaid.js from DLL resource to local directory
    ExtractMermaidJs();

    WCHAR userDataDir[MAX_PATH] = {};
    if (SUCCEEDED(SHGetFolderPathW(nullptr, CSIDL_LOCAL_APPDATA, nullptr, 0, userDataDir))) {
        wcscat_s(userDataDir, L"\\MermaidPreview\\WebView2");
    }

    HRESULT hr = CreateCoreWebView2EnvironmentWithOptions(
        nullptr, userDataDir, nullptr,
        Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>(
            [this](HRESULT result, ICoreWebView2Environment* env) -> HRESULT {
                if (m_bDestroyed) return E_ABORT;
                if (FAILED(result) || !env)
                    return result;

                m_env = env;

                return env->CreateCoreWebView2Controller(
                    m_hwndParent,
                    Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
                        [this](HRESULT result, ICoreWebView2Controller* controller) -> HRESULT {
                            if (m_bDestroyed) return E_ABORT;
                            if (FAILED(result) || !controller)
                                return result;

                            m_controller = controller;
                            HRESULT hrGet = controller->get_CoreWebView2(&m_webview);
                            if (FAILED(hrGet) || !m_webview)
                                return FAILED(hrGet) ? hrGet : E_FAIL;

                            ComPtr<ICoreWebView2Settings> settings;
                            m_webview->get_Settings(&settings);
                            settings->put_IsScriptEnabled(TRUE);
                            settings->put_AreDefaultScriptDialogsEnabled(FALSE);
                            settings->put_IsWebMessageEnabled(TRUE);
                            settings->put_AreDevToolsEnabled(FALSE);
                            settings->put_IsStatusBarEnabled(FALSE);
                            settings->put_AreDefaultContextMenusEnabled(FALSE);
                            settings->put_AreHostObjectsAllowed(FALSE);

                            RECT bounds;
                            GetClientRect(m_hwndParent, &bounds);
                            m_controller->put_Bounds(bounds);

                            // Write HTML page to local directory and navigate to it
                            // This allows mermaid.js to load from the same directory
                            std::wstring htmlPath = m_resourceDir + L"\\preview.html";

                            // Optimization 1: Check if existing preview.html has correct version
                            bool needWrite = true;
                            {
                                HANDLE hCheck = CreateFileW(htmlPath.c_str(), GENERIC_READ,
                                    FILE_SHARE_READ, nullptr, OPEN_EXISTING,
                                    FILE_ATTRIBUTE_NORMAL, nullptr);
                                if (hCheck != INVALID_HANDLE_VALUE) {
                                    char buf[64] = {};
                                    DWORD bytesRead = 0;
                                    ReadFile(hCheck, buf, sizeof(buf) - 1, &bytesRead, nullptr);
                                    CloseHandle(hCheck);
                                    // Skip UTF-8 BOM (3 bytes) if present
                                    const char* p = (bytesRead >= 3 &&
                                        (BYTE)buf[0] == 0xEF && (BYTE)buf[1] == 0xBB &&
                                        (BYTE)buf[2] == 0xBF) ? buf + 3 : buf;
                                    if (strstr(p, kHtmlVersionTag))
                                        needWrite = false;
                                }
                            }

                            if (needWrite) {
                                std::wstring htmlContent = BuildHtmlPage();

                                // Convert wstring to UTF-8 for file writing
                                int utf8Len = WideCharToMultiByte(CP_UTF8, 0, htmlContent.c_str(),
                                    (int)htmlContent.size(), nullptr, 0, nullptr, nullptr);
                                std::string utf8Content(utf8Len, '\0');
                                WideCharToMultiByte(CP_UTF8, 0, htmlContent.c_str(),
                                    (int)htmlContent.size(), utf8Content.data(), utf8Len, nullptr, nullptr);

                                HANDLE hFile = CreateFileW(htmlPath.c_str(), GENERIC_WRITE, 0,
                                    nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
                                if (hFile != INVALID_HANDLE_VALUE) {
                                    // Write UTF-8 BOM
                                    const BYTE bom[] = { 0xEF, 0xBB, 0xBF };
                                    DWORD written = 0;
                                    WriteFile(hFile, bom, 3, &written, nullptr);
                                    WriteFile(hFile, utf8Content.data(), (DWORD)utf8Content.size(), &written, nullptr);
                                    CloseHandle(hFile);
                                }
                            }

                            // Navigate to local file
                            std::wstring fileUrl = L"file:///" + m_resourceDir + L"\\preview.html";
                            // Replace backslashes with forward slashes for URL
                            for (auto& c : fileUrl) { if (c == L'\\') c = L'/'; }
                            m_webview->Navigate(fileUrl.c_str());

                            m_webview->add_NavigationCompleted(
                                Callback<ICoreWebView2NavigationCompletedEventHandler>(
                                    [this](ICoreWebView2* /*sender*/,
                                           ICoreWebView2NavigationCompletedEventArgs* args) -> HRESULT {
                                        if (m_bDestroyed) return S_OK;
                                        BOOL success = FALSE;
                                        args->get_IsSuccess(&success);
                                        if (success) {
                                            m_bReady = true;
                                            if (m_onReady)
                                                m_onReady();

                                            if (m_hasPendingRender) {
                                                RenderContent(m_pendingHtml, m_pendingDarkMode);
                                                m_pendingHtml.clear();
                                                m_hasPendingRender = false;
                                            }
                                        }
                                        return S_OK;
                                    }).Get(),
                                &m_navCompletedToken);

                            return S_OK;
                        }).Get());
            }).Get());

    return hr;
}

// ============================================================================
// RenderContent - Send pre-parsed HTML to WebView2
// ============================================================================
void WebView2Manager::RenderContent(const std::wstring& htmlContent, bool darkMode)
{
    if (!m_bReady || !m_webview) {
        m_pendingHtml = htmlContent;
        m_pendingDarkMode = darkMode;
        m_hasPendingRender = true;
        return;
    }

    // Build JS call: renderContent('...escaped HTML...', 'dark'|'light')
    std::wstring js = L"renderContent('";
    js += EscapeForJS(htmlContent);
    js += L"', '";
    js += darkMode ? L"dark" : L"light";
    js += L"');";

    m_webview->ExecuteScript(js.c_str(), nullptr);
}

void WebView2Manager::SetTheme(bool darkMode)
{
    if (!m_bReady || !m_webview)
        return;

    std::wstring js = darkMode ? L"setTheme(true);" : L"setTheme(false);";
    m_webview->ExecuteScript(js.c_str(), nullptr);
}

void WebView2Manager::Resize(RECT bounds)
{
    if (m_controller) {
        m_controller->put_Bounds(bounds);
    }
}

void WebView2Manager::Clear()
{
    if (!m_bReady || !m_webview)
        return;

    m_webview->ExecuteScript(L"clearContent();", nullptr);
}

// ============================================================================
// SetEditCallback - Register WebMessage handler for editing from preview
// ============================================================================
void WebView2Manager::SetEditCallback(EditCallback callback)
{
    m_editCallback = std::move(callback);

    if (!m_webview || m_bMessageHandlerRegistered)
        return;

    m_bMessageHandlerRegistered = true;

    m_webview->add_WebMessageReceived(
        Callback<ICoreWebView2WebMessageReceivedEventHandler>(
            [this](ICoreWebView2* /*sender*/,
                   ICoreWebView2WebMessageReceivedEventArgs* args) -> HRESULT {
                LPWSTR jsonRaw = nullptr;
                args->get_WebMessageAsJson(&jsonRaw);
                if (!jsonRaw) return S_OK;

                std::wstring json = jsonRaw;
                CoTaskMemFree(jsonRaw);

                // Extract the canonical "type" field once and dispatch on
                // exact equality. Substring search ("does the JSON contain
                // the literal word 'theme'?") is fooled by any payload that
                // happens to mention the keyword, e.g. an `edit` message
                // whose newText contains the word "theme".
                std::wstring msgType;
                {
                    size_t kp = json.find(L"\"type\"");
                    if (kp == std::wstring::npos) return S_OK;
                    kp += 6;
                    // Skip ':' and whitespace
                    while (kp < json.size() && (json[kp] == L':' || json[kp] == L' '))
                        kp++;
                    if (kp >= json.size() || json[kp] != L'"') return S_OK;
                    kp++;
                    // Read until closing quote (no escape handling needed —
                    // the JS senders only emit ASCII identifiers as type).
                    size_t kpEnd = kp;
                    while (kpEnd < json.size() && json[kpEnd] != L'"' &&
                           kpEnd - kp < 32)
                        kpEnd++;
                    if (kpEnd >= json.size() || json[kpEnd] != L'"') return S_OK;
                    msgType = json.substr(kp, kpEnd - kp);
                }
                if (msgType.empty()) return S_OK;

                // Dispatch: theme change message
                if (msgType == L"theme") {
                    if (m_themeCallback) {
                        bool dark = json.find(L"\"dark\":true") != std::wstring::npos;
                        m_themeCallback(dark);
                    }
                    return S_OK;
                }

                // Safe integer extraction with overflow protection
                constexpr int MAX_LINE_NUMBER = 10000000;
                auto safeExtractLine = [&](const std::wstring& key) -> int {
                    size_t pos = json.find(key);
                    if (pos == std::wstring::npos) return -1;
                    pos += key.size();
                    while (pos < json.size() && (json[pos] == L':' || json[pos] == L' '))
                        pos++;
                    int val = 0;
                    bool found = false;
                    while (pos < json.size() && json[pos] >= L'0' && json[pos] <= L'9') {
                        int digit = json[pos] - L'0';
                        if (val > (INT_MAX - digit) / 10)
                            return -1; // Overflow
                        val = val * 10 + digit;
                        pos++;
                        found = true;
                    }
                    if (!found || val > MAX_LINE_NUMBER) return -1;
                    return val;
                };

                // Dispatch: fontSize message
                if (msgType == L"fontSize") {
                    if (m_fontSizeCallback) {
                        int val = safeExtractLine(L"\"size\"");
                        if (val >= 8 && val <= 32) m_fontSizeCallback(val);
                    }
                    return S_OK;
                }

                // Dispatch: syncScroll message (Preview → Editor)
                if (msgType == L"syncScroll") {
                    if (m_scrollCallback) {
                        int val = safeExtractLine(L"\"line\"");
                        if (val >= 0) m_scrollCallback(val);
                    }
                    return S_OK;
                }

                // Dispatch: navigateToLine message (click mermaid → jump editor)
                if (msgType == L"navigateToLine") {
                    if (m_navigateCallback) {
                        int val = safeExtractLine(L"\"line\"");
                        if (val >= 0) m_navigateCallback(val);
                    }
                    return S_OK;
                }

                // Dispatch: openFile message (relative path link clicked)
                if (msgType == L"openFile" && m_openFileCallback) {
                    // Extract "path" string value
                    std::wstring filePath;
                    size_t pathPos = json.find(L"\"path\"");
                    if (pathPos != std::wstring::npos) {
                        pathPos = json.find(L'"', pathPos + 6);
                        if (pathPos != std::wstring::npos) {
                            pathPos++;
                            while (pathPos < json.size() && json[pathPos] != L'"') {
                                if (json[pathPos] == L'\\' && pathPos + 1 < json.size()) {
                                    switch (json[pathPos + 1]) {
                                    case L'"':  filePath += L'"';  pathPos += 2; break;
                                    case L'\\': filePath += L'\\'; pathPos += 2; break;
                                    case L'/':  filePath += L'/';  pathPos += 2; break;
                                    default:    filePath += json[pathPos]; pathPos++; break;
                                    }
                                } else {
                                    filePath += json[pathPos];
                                    pathPos++;
                                }
                            }
                        }
                    }
                    if (!filePath.empty()) {
                        m_openFileCallback(filePath);
                    }
                    return S_OK;
                }

                // Dispatch: editMermaidNode (M2 — inline mermaid label edit).
                // Inputs: blockId="mermaid-placeholder-N", nodeId, newLabel
                // (newLabel arrives with `\n` representing `<br/>` in source).
                if (msgType == L"editMermaidNode" && m_editMermaidNodeCallback) {
                    auto extractStr = [&](const std::wstring& key) -> std::wstring {
                        std::wstring v;
                        size_t kp = json.find(key);
                        if (kp == std::wstring::npos) return v;
                        kp = json.find(L'"', kp + key.size());
                        if (kp == std::wstring::npos) return v;
                        ++kp;
                        while (kp < json.size() && json[kp] != L'"') {
                            if (json[kp] == L'\\' && kp + 1 < json.size()) {
                                switch (json[kp + 1]) {
                                case L'n':  v += L'\n'; kp += 2; break;
                                case L'r':  v += L'\r'; kp += 2; break;
                                case L't':  v += L'\t'; kp += 2; break;
                                case L'"':  v += L'"';  kp += 2; break;
                                case L'\\': v += L'\\'; kp += 2; break;
                                case L'/':  v += L'/';  kp += 2; break;
                                default:    v += json[kp]; ++kp; break;
                                }
                            } else { v += json[kp]; ++kp; }
                        }
                        return v;
                    };
                    std::wstring blockId  = extractStr(L"\"blockId\"");
                    std::wstring nodeId   = extractStr(L"\"nodeId\"");
                    std::wstring newLabel = extractStr(L"\"newLabel\"");

                    // Validate blockId == "mermaid-placeholder-<digits>".
                    bool blockOk = blockId.rfind(L"mermaid-placeholder-", 0) == 0;
                    if (blockOk) {
                        for (size_t bi = 20; bi < blockId.size(); ++bi) {
                            if (blockId[bi] < L'0' || blockId[bi] > L'9') { blockOk = false; break; }
                        }
                        if (blockId.size() <= 20) blockOk = false;
                    }
                    // Validate nodeId — mermaid identifier grammar.
                    bool nodeOk = !nodeId.empty();
                    for (wchar_t c : nodeId) {
                        if (!((c >= L'A' && c <= L'Z') || (c >= L'a' && c <= L'z') ||
                              (c >= L'0' && c <= L'9') || c == L'_' || c == L'-')) {
                            nodeOk = false; break;
                        }
                    }
                    // Reuse the existing 1 MB cap from the plain-text edit path.
                    if (blockOk && nodeOk && newLabel.size() <= 1024 * 1024) {
                        m_editMermaidNodeCallback(blockId, nodeId, newLabel);
                    }
                    return S_OK;
                }

                // Dispatch: edit message
                if (msgType != L"edit" || !m_editCallback)
                    return S_OK;

                // Parse lineStart and lineEnd (reuse safeExtractLine)
                int lineStart = safeExtractLine(L"\"lineStart\"");
                int lineEnd = safeExtractLine(L"\"lineEnd\"");

                // Extract newText string value
                std::wstring newText;
                size_t textPos = json.find(L"\"newText\"");
                if (textPos != std::wstring::npos) {
                    textPos = json.find(L'"', textPos + 9);
                    if (textPos != std::wstring::npos) {
                        textPos++;
                        while (textPos < json.size() && json[textPos] != L'"') {
                            if (json[textPos] == L'\\' && textPos + 1 < json.size()) {
                                switch (json[textPos + 1]) {
                                case L'n':  newText += L'\n'; textPos += 2; break;
                                case L'r':  newText += L'\r'; textPos += 2; break;
                                case L't':  newText += L'\t'; textPos += 2; break;
                                case L'"':  newText += L'"';  textPos += 2; break;
                                case L'\\': newText += L'\\'; textPos += 2; break;
                                default:    newText += json[textPos]; textPos++; break;
                                }
                            } else {
                                newText += json[textPos];
                                textPos++;
                            }
                        }
                    }
                }

                if (lineStart >= 0 && lineEnd >= 0) {
                    m_editCallback(lineStart, lineEnd, newText);
                }

                return S_OK;
            }).Get(),
        &m_webMessageToken);
}

// ============================================================================
// SetEditMermaidNodeCallback
// ============================================================================
void WebView2Manager::SetEditMermaidNodeCallback(EditMermaidNodeCallback callback)
{
    m_editMermaidNodeCallback = std::move(callback);
}

// ============================================================================
// SetThemeCallback
// ============================================================================
void WebView2Manager::SetThemeCallback(ThemeCallback callback)
{
    m_themeCallback = std::move(callback);
}

// ============================================================================
// ScrollToLine - Scroll preview to match editor's top visible line
// ============================================================================
void WebView2Manager::ScrollToLine(int line)
{
    if (!m_bReady || !m_webview)
        return;

    std::wstring js = L"scrollToLine(" + std::to_wstring(line) + L");";
    m_webview->ExecuteScript(js.c_str(), nullptr);
}

// ============================================================================
// SetScrollCallback
// ============================================================================
void WebView2Manager::SetScrollCallback(ScrollCallback callback)
{
    m_scrollCallback = std::move(callback);
}

void WebView2Manager::SetNavigateCallback(NavigateCallback callback)
{
    m_navigateCallback = std::move(callback);
}

void WebView2Manager::SetOpenFileCallback(OpenFileCallback callback)
{
    m_openFileCallback = std::move(callback);
}

void WebView2Manager::SetFontSizeCallback(FontSizeCallback callback)
{
    m_fontSizeCallback = std::move(callback);
}

void WebView2Manager::SetFontSize(int size)
{
    if (!m_bReady || !m_webview)
        return;

    if (size < 8) size = 8;
    if (size > 32) size = 32;

    std::wstring js = L"setFontSize(" + std::to_wstring(size) + L");";
    m_webview->ExecuteScript(js.c_str(), nullptr);
}

// ============================================================================
// Park - Move WebView2 to a hidden parking window without destroying
// ============================================================================
void WebView2Manager::Park(HWND hwndParking)
{
    if (!m_controller)
        return;

    // Hide first to deactivate WebView2 and release focus,
    // THEN reparent to parking window.
    m_controller->put_IsVisible(FALSE);
    m_controller->put_ParentWindow(hwndParking);
    m_hwndParent = hwndParking;
}

// ============================================================================
// Reparent - Move WebView2 from parking window to a new host (fast reopen)
// ============================================================================
HRESULT WebView2Manager::Reparent(HWND hwndNewParent)
{
    if (!m_controller)
        return E_FAIL;

    HRESULT hr = m_controller->put_ParentWindow(hwndNewParent);
    if (FAILED(hr))
        return hr;

    m_hwndParent = hwndNewParent;

    RECT bounds;
    GetClientRect(hwndNewParent, &bounds);
    m_controller->put_Bounds(bounds);
    m_controller->put_IsVisible(TRUE);

    return S_OK;
}

void WebView2Manager::Destroy()
{
    m_bDestroyed = true;
    m_bReady = false;
    m_hasPendingRender = false;
    m_pendingHtml.clear();
    m_bMessageHandlerRegistered = false;

    // Unsubscribe COM event handlers before releasing objects
    if (m_webview) {
        m_webview->remove_NavigationCompleted(m_navCompletedToken);
        m_webview->remove_WebMessageReceived(m_webMessageToken);
        m_navCompletedToken = {};
        m_webMessageToken = {};
    }

    if (m_controller) {
        m_controller->Close();
        m_controller.Reset();
    }
    m_webview.Reset();
    m_env.Reset();
}
