#include "WebView2Manager.h"
#include "resource.h"
#include <shlobj.h>
#include <sstream>
#include <climits>

using namespace Microsoft::WRL;

// Get the DLL's HMODULE (for loading resources)
extern HINSTANCE EEGetInstanceHandle();

// HTML cache version tag — increment when BuildHtmlPage() content changes
static const char* kHtmlVersionTag = "<!-- MermaidPreview-v11 -->";

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

    // Write to file
    HANDLE hFile = CreateFileW(mermaidPath.c_str(), GENERIC_WRITE, 0, nullptr,
                               CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (hFile == INVALID_HANDLE_VALUE)
        return false;

    DWORD written = 0;
    WriteFile(hFile, data, size, &written, nullptr);
    CloseHandle(hFile);

    return written == size;
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
    std::wstring html = LR"P1(<!-- MermaidPreview-v11 -->
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
  .mermaid-container { margin: 1em 0; overflow: hidden; cursor: grab; text-align: center; position: relative; border-radius: 4px; min-height: 40px; }
  .mermaid-container.dragging { cursor: grabbing; }
  .mermaid-container.zoom-active { outline: 2px dashed #0969da; outline-offset: 2px; cursor: zoom-in; }
  body.dark .mermaid-container.zoom-active { outline-color: #58a6ff; }
  .mermaid-container.zoom-active .svg-zoom-badge { opacity:1 !important; background:rgba(9,105,218,0.7); }
  body.dark .mermaid-container.zoom-active .svg-zoom-badge { background:rgba(88,166,255,0.7); }
  .mermaid-container svg { display: inline-block; max-width: 100%; height: auto; transform-origin: 0 0; user-select: none; }
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
  </div>
)P1a";

    // --- Part 1b: Script opening + mermaid loader ---
    html += LR"P1b(  <script>
    var mermaidReady = false;
    var renderedMermaidSrcs = {};
    var zoomTarget = null;
    var _pendingRender = null;

    // Async load mermaid.min.js — does not block NavigationCompleted
    (function(){
      var s = document.createElement('script');
      s.src = ')P1b" + mermaidSrc + LR"P2(';
      s.onload = function() {
        mermaid.initialize({ startOnLoad:false, theme:'default', securityLevel:'strict', flowchart:{useMaxWidth:true}, sequence:{useMaxWidth:true} });
        mermaidReady = true;
        if (_pendingRender) { _processPendingMermaid(); }
      };
      document.head.appendChild(s);
    })();

    async function _processPendingMermaid() {
      if (!_pendingRender) return;
      var isDark = (_pendingRender.theme === 'dark');
      mermaid.initialize({ startOnLoad:false, theme: isDark?'dark':'default', securityLevel:'strict', flowchart:{useMaxWidth:true}, sequence:{useMaxWidth:true} });
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
      container.querySelectorAll('.mermaid-container').forEach(initSvgDrag);
    }

    window.renderContent = async function(htmlContent, theme) {
      var isDark = (theme === 'dark');
      document.body.className = isDark ? 'dark' : 'light';
      if (mermaidReady) {
        mermaid.initialize({ startOnLoad:false, theme: isDark?'dark':'default', securityLevel:'strict', flowchart:{useMaxWidth:true}, sequence:{useMaxWidth:true} });
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
      container.querySelectorAll('.mermaid-container').forEach(initSvgDrag);
    };
    window.setTheme = function(dark) { document.body.className = dark ? 'dark' : 'light'; };
    window.clearContent = function() { document.getElementById('content').innerHTML = '<div class="empty">Open a Markdown file to preview</div>'; renderedMermaidSrcs = {}; _pendingRender = null; };
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
    document.getElementById('content').addEventListener('dblclick', function(e) {
      var el = e.target;
      while (el && el !== this) { if (el.hasAttribute('data-line-start')) break; el = el.parentElement; }
      if (!el || !el.hasAttribute('data-line-start')) return;
      if (el.closest('.mermaid-container')) return;
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
      // Handle internal anchor links (#section)
      var anchor = e.target.closest('a[href]');
      if(anchor){
        var href = anchor.getAttribute('href');
        if(href && href.charAt(0)==='#'){
          e.preventDefault();
          var target = document.getElementById(href.substring(1));
          if(target) target.scrollIntoView({behavior:'smooth',block:'start'});
        }
      }
    });
    window.addEventListener('blur', function(){ ctxMenu.style.display='none'; });
    document.querySelectorAll('.ctx-theme').forEach(function(el){ el.addEventListener('click',function(){ switchTheme(this.dataset.dark==='true'); ctxMenu.style.display='none'; }); });
    document.querySelectorAll('.ctx-action').forEach(function(el){ el.addEventListener('click',function(){
      var a=this.dataset.action;
      if(a==='zoom-in') svgZoomAll(1.25); else if(a==='zoom-out') svgZoomAll(0.8); else if(a==='reset') svgResetAll();
      ctxMenu.style.display='none';
    }); });

    function switchTheme(dark) {
      document.body.className = dark ? 'dark' : 'light';
      if (mermaidReady) {
        mermaid.initialize({startOnLoad:false, theme:dark?'dark':'default', securityLevel:'strict', flowchart:{useMaxWidth:true}, sequence:{useMaxWidth:true}});
        document.querySelectorAll('.mermaid-container[data-mermaid-src]').forEach(function(el,idx){
          var src = decodeURIComponent(el.getAttribute('data-mermaid-src'));
          mermaid.render('ts-'+idx+'-'+Date.now(), src).then(function(r){ el.innerHTML=r.svg; initSvgDrag(el); }).catch(function(){});
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
                            controller->get_CoreWebView2(&m_webview);

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

                if (json.find(L"\"type\"") == std::wstring::npos)
                    return S_OK;

                // Dispatch: theme change message
                if (json.find(L"\"theme\"") != std::wstring::npos && m_themeCallback) {
                    bool dark = json.find(L"\"dark\":true") != std::wstring::npos;
                    m_themeCallback(dark);
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

                // Dispatch: syncScroll message (Preview → Editor)
                if (json.find(L"\"syncScroll\"") != std::wstring::npos && m_scrollCallback) {
                    int val = safeExtractLine(L"\"line\"");
                    if (val >= 0) m_scrollCallback(val);
                    return S_OK;
                }

                // Dispatch: navigateToLine message (click mermaid → jump editor)
                if (json.find(L"\"navigateToLine\"") != std::wstring::npos && m_navigateCallback) {
                    int val = safeExtractLine(L"\"line\"");
                    if (val >= 0) m_navigateCallback(val);
                    return S_OK;
                }

                // Dispatch: edit message
                if (json.find(L"\"edit\"") == std::wstring::npos || !m_editCallback)
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
