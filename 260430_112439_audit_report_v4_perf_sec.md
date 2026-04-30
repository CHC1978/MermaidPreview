# Code Audit Report v4 — Performance & Security
**Generated:** 260430_112439
**Reviewer:** qa-agent (code-audit-optimizer)
**Tech Stack:** C++17 / MSVC, WebView2, JavaScript (inline HTML shell), Win32
**Scope:** 新增 M1/M2/M3/P1.4-P1.7 功能 — WebView2Manager.cpp JS 端、MarkdownParser ExtractFlowchartNodes、MermaidPreview.cpp async render 路徑

---

## Executive Summary

| 優先級 | 數量 |
|--------|------|
| Critical | 2 |
| High | 4 |
| Medium | 5 |
| Low | 3 |

**Top-3 優先處理:**
1. **[SEC-001] `probe.innerHTML = inner.innerHTML`** — XSS 注入，cluster label 原始 SVG 文字未清理直接塞入 HTML 測量容器
2. **[SEC-002] `_wrapLongSubgraphTitles` `<br/>` 注入到 mermaid 原始碼** — 攻擊者可利用 cluster title 注入任意 mermaid 指令
3. **[PERF-001] `_detectOverlaps` N×M getBoundingClientRect 每次 render 觸發強制 reflow** — 50+ 節點圖 UI blocking ~200ms

---

## 1. 安全性漏洞 (Security Vulnerabilities)

### SEC-001 [Critical] `_growClusterLabels` — innerHTML XSS 注入

**位置:** `src/WebView2Manager.cpp`, BuildHtmlPage Part1 `_growClusterLabels` 函式, 第 358–359 行（相對於 `html +=` 段落）

```javascript
// 問題程式碼
var probe = document.createElement('div');
probe.style.cssText = 'display:inline-block;white-space:nowrap;';
probe.innerHTML = inner.innerHTML;   // ← 直接複製 SVG 的 innerHTML
measurer.appendChild(probe);
```

**根因:**
`inner` 是 mermaid SVG 內 `<foreignObject>` 裡的 `<div>`，其 `innerHTML` 來自 mermaid 渲染輸出。雖然 mermaid 使用 `securityLevel: 'strict'`，但 strict 模式只會 sanitize 輸入的 mermaid *語法*，不會清理已渲染 SVG 中的 HTML 內容。如果 mermaid 11.14 任何版本在 cluster title 渲染中存在 XSS bypass（已有歷史先例），`inner.innerHTML` 可攜帶惡意 HTML 並被注入到 `document.body` 的測量 div 中。

**影響:**
WebView2 在 `securityLevel: 'strict'` 下本身有 DOMPurify 防護，但 `probe.innerHTML` 直接繞過了任何 sanitization 層，屬於 defense-in-depth 缺口。若 mermaid 本身被旁路，可觸發 WebMessage 欺騙、exfiltrate cluster label 內容。

**建議修復:**
```javascript
// 改用 textContent 複製（純文字），或先通過 DOMPurify
probe.textContent = inner.textContent;
// 若需要保留 <span> 樣式，使用 DOMPurify（mermaid 本身就內建）:
// probe.innerHTML = DOMPurify.sanitize(inner.innerHTML, {ALLOWED_TAGS:['span','br']});
```

---

### SEC-002 [Critical] `_wrapLongSubgraphTitles` — Mermaid 原始碼注入

**位置:** `src/WebView2Manager.cpp`, `_wrapLongSubgraphTitles` 函式

```javascript
function _wrapLongSubgraphTitles(src) {
  var rx = /^(\s*subgraph\s+\S+\s*\[\s*")([^"]+)("\s*\])\s*$/gm;
  return src.replace(rx, function(whole, head, title, tail) {
    if (/<br\s*\/?>/i.test(title)) return whole;
    if (title.length <= 28) return whole;
    // ...
    return head + left + '<br/>' + right + tail;  // ← 拼接後送回 C++
  });
}
```

**根因:**
此函式對從 `data-mermaid-src` attribute 讀取的 mermaid 原始碼做字串替換，並將結果透過 `postMessage({type:'editMermaidBlock', newSource})` 送回 C++，再寫回 EmEditor 文件。`<br/>` 是硬碼注入到 `title` 中的，但 `title` 本身（`[^"]+`）未做任何清理。

攻擊向量：若 mermaid 文件由惡意來源提供（e-mail 附件、共用文件），且 cluster title 包含精心設計的字串（例如內含 `%% inject` 的偽裝標題），`_wrapLongSubgraphTitles` 可能在換行分割時造成非預期的 mermaid 語法插入。更直接的問題是：`<br/>` 被注入為字面字串進入 mermaid 原始碼，而 mermaid 11 在某些場合會把 `<br/>` 解析為 HTML，可能導致最終渲染的 SVG 包含 HTML 注入節點。

**影響:** High，攻擊者控制文件後可透過 auto-fix 路徑修改其他使用者的文件（如果文件是協作共享的）。

**建議修復:**
1. `<br/>` 注入應改為 mermaid 官方支援的換行方法（若 mermaid 版本不支援 br，改用空格截斷）。
2. 對 `title` 做二次 regex 確認，拒絕包含 `%%`、`\n`、`-->` 的 title。

---

### SEC-003 [High] `editMermaidNode` WebMessage `nodeId` 驗證強度不足

**位置:** `src/WebView2Manager.cpp`, `SetEditCallback` lambda, 第 1460–1465 行

```cpp
bool nodeOk = !nodeId.empty();
for (wchar_t c : nodeId) {
    if (!((c >= L'A' && c <= L'Z') || (c >= L'a' && c <= L'z') ||
          (c >= L'0' && c <= L'9') || c == L'_' || c == L'-')) {
        nodeOk = false; break;
    }
}
```

JS 端在 `rawId.match(/-flowchart-(.+)-\d+$/)` 提取 nodeId 時使用了 `.+`（貪婪任意字元）。若 SVG id 格式被 mermaid 未來版本更改，capture group 可能包含非 ASCII 字元或特殊符號（如 `.`、`/`）。

**根因:** C++ 端雖有 ASCII-only 白名單，但 nodeId 長度上限缺失（空字串已擋，但超長字串未限制）。配合 `extractStr` 的簡易 JSON 解析器，若 `newLabel` 欄位 value 中包含 `"nodeId":"` 子字串，可能觸發欄位提取錯位（JSON key 衝突）。

**建議修復:**
```cpp
// 加上長度上限
if (nodeId.size() > 128) nodeOk = false;
```

---

### SEC-004 [High] `extractStrField` / `extractStr` 手寫 JSON 解析器 — Unicode escape 未處理

**位置:** `src/WebView2Manager.cpp`, `editMermaidBlock` 與 `editMermaidNode` 分支，兩個幾乎相同的 lambda

```cpp
// 只處理了 \n \r \t \" \\ \/
// 未處理 \uXXXX Unicode escape
switch (json[kp + 1]) {
case L'n':  v += L'\n'; kp += 2; break;
case L'r':  v += L'\r'; kp += 2; break;
// ...
default:    v += json[kp]; ++kp; break;  // ← \uXXXX 被當作字面 \u
}
```

**根因:** JSON 規範要求處理 `\uXXXX` escape。若 JS 端 `JSON.stringify` 了包含 CJK 字元或 emoji 的 mermaid label（`newLabel`），某些字元會被 escape 為 `中`。C++ 端的 `default` 分支會輸出字面 `\` + `u` 到文字，導致 EmEditor 寫回的文件含有損壞的字元 escape。

**影響:** Medium-High — 使用者編輯含 CJK 字元的 mermaid node label 後，文件被寫入損壞的字串（靜默資料損失）。

**建議修復:**
```cpp
case L'u': {
    // 解析 \uXXXX
    if (kp + 5 < json.size()) {
        wchar_t code = 0;
        for (int di = 2; di <= 5; ++di) {
            wchar_t h = json[kp + di];
            int val = (h >= L'0' && h <= L'9') ? h - L'0'
                    : (h >= L'a' && h <= L'f') ? h - L'a' + 10
                    : (h >= L'A' && h <= L'F') ? h - L'A' + 10 : -1;
            if (val < 0) { code = 0; break; }
            code = (wchar_t)((code << 4) | val);
        }
        if (code) { v += code; kp += 6; break; }
    }
    v += json[kp]; ++kp; // fallback
    break;
}
```

---

### SEC-005 [High] `_applyAutoFix` — `data-mermaid-src` decodeURIComponent 未清理後直接進入 `_injectSpacingDirective` regex 替換

**位置:** `src/WebView2Manager.cpp`, `_applyAutoFix` 函式

```javascript
var rawSrc = container.getAttribute('data-mermaid-src');
// ...
try { src = decodeURIComponent(rawSrc); }
catch(_) { _showToast(container, 'Cannot decode source', true); return; }

var newSrc = src;
// ...直接進入 _injectSpacingDirective(newSrc)
```

**根因:** `data-mermaid-src` 由 C++ 端 `UrlEncode` 編碼後塞入 HTML attribute，理論上應為安全的。但 C++ 端的 `UrlEncode` 保留了空格（轉為 `%20`），而 decodeURIComponent 後的字串未再次驗證是否含有 `\n`（newline）或 `%%{`（mermaid directive）前綴。如果 mermaid 原始碼含有特定序列（例如 `%%{init:...}%%` 後接 rankSpacing），`_injectSpacingDirective` 的 regex 會在第一個 `%%{...init...}%%` directive 找到後更新，但若 doc 中有多個 directive，只有第一個被 regex `/%%\{[^%]*init[^%]*\}%%/.test(src)` 匹配，第二個 `rankSpacing` 值保持不動，產生意外的 spacing 設定。

**影響:** Low-Medium（功能性錯誤，非直接 injection），但值得修正以避免 auto-fix 靜默失效。

---

### SEC-006 [Medium] `OnPreviewMermaidBlockEdited` — newSource 未驗證 mermaid fence 語法

**位置:** `src/MermaidPreview.cpp`, `OnPreviewMermaidBlockEdited`, 第 1242–1301 行

JS 端雖然呼叫了 `mermaid.parse(newSrc)` 驗證，但 `mermaid.parse` 對 `_injectSpacingDirective` 產生的 `%%{init:...}%%` 指令可能不會報錯，即使指令格式略有損壞。C++ 端收到 `newSource` 後直接重組：

```cpp
newBlock += openFence;
newBlock += L'\n';
newBlock += body;      // ← body 來自 JS，未做二次 C++ 端驗證
newBlock += L'\n';
newBlock += closeFence;
```

**影響:** 已有 4MB cap 防止超大 payload，但無防止 body 內含 ` ``` ` 字串（會提前關閉 fenced code block）。這是低風險，因為 `mermaid.parse` 通常會拒絕含有 ` ``` ` 的字串，但仍屬防禦深度不足。

**建議修復:** C++ 端加一行：
```cpp
if (body.find(L"```") != std::wstring::npos) return;
```

---

## 2. 效能瓶頸 (Performance Bottlenecks)

### PERF-001 [High] `_detectOverlaps` — 強制 reflow N×M AABB 比對，每次 render 後執行

**位置:** `src/WebView2Manager.cpp`, `_detectOverlaps` + `_refreshAutoFixBtn`

```javascript
function _detectOverlaps(container) {
  var nodes = svg.querySelectorAll('g.node');
  var edgeLabels = svg.querySelectorAll('g.edgeLabels g.label');
  var clusterLabels = svg.querySelectorAll('g.cluster-label');
  var nodeBoxes = [];
  for (var i = 0; i < nodes.length; i++) {
    var nb = nodes[i].getBoundingClientRect();   // ← 每個 node 各觸發 reflow
    ...
  }
  // probe() 再各自 getBoundingClientRect()
}
```

**根因:**
- `getBoundingClientRect()` 是強制 layout flush 操作（forced reflow）。
- 對 N nodes + M edge labels + K cluster labels，每次呼叫均觸發 style recalc + layout。
- 函式在每次 render 後被每個 container 各呼叫一次。
- 已有 `requestAnimationFrame` 延遲（`_refreshAutoFixBtn` 使用 rAF），這只減少了「阻塞主線程的時間點」，不減少 reflow 次數。

**量化影響:**
- 14-node 圖：14 + (edge labels N) + (cluster labels K) 次 `getBoundingClientRect` = ~20–30 次強制 reflow。
- 50-node RE 圖：估計 50–80 次，在 WebView2 的 Chromium 引擎下，每次 reflow ~1–5ms，總計 50–400ms UI blocking。

**建議優化（短期）:**
```javascript
// Batch：先讀完所有 rect，再做比對
var nodeBoxes = Array.from(nodes).map(function(n) { return n.getBoundingClientRect(); });
var edgeBoxes = Array.from(edgeLabels).map(function(l) { return l.getBoundingClientRect(); });
var clusterBoxes = Array.from(clusterLabels).map(function(l) { return l.getBoundingClientRect(); });
// 上面三行是連續讀取，瀏覽器可能合併 layout flush
// 之後純做 JS 計算，不再觸發 reflow
```

**建議優化（長期）:**
- 使用 `IntersectionObserver` 替代主動 polling（WebView2/Chromium 完整支援）。
- 對大型圖（`nodes.length > 30`），限制 overlap 偵測 sample 比例。

---

### PERF-002 [High] `_growClusterLabels` — 每個 cluster label 各插入/移除測量 div，觸發多次 reflow

**位置:** `src/WebView2Manager.cpp`, `_growClusterLabels` 函式

```javascript
document.body.appendChild(measurer);  // 1 次 DOM 插入（reflow）
try {
  for (var i = 0; i < fos.length; i++) {
    measurer.innerHTML = '';            // ← 每 iteration 清空（可能 reflow）
    // ...
    probe.offsetWidth || probe.scrollWidth  // ← 每 iteration 強制 reflow
  }
} finally {
  measurer.parentNode.removeChild(measurer); // 移除
}
```

**根因:**
- `measurer.innerHTML = ''` 在每個 cluster 都清空並重建 `probe` 子元素。
- `probe.offsetWidth` 是強制 layout flush。
- 對 K 個 cluster labels = K 次 reflow（在 measurer 插入 body 之後的每輪）。
- 每次 render 後的 `_liftLabels` 會呼叫 `_growClusterLabels`。

**量化影響:** K cluster × ~2–5ms/reflow = 對 10 cluster 的圖約 20–50ms。

**建議優化:**
```javascript
// 批次：一次把所有 probe 都 append 進 measurer，一次 reflow，讀完再清
var measurer = document.createElement('div');
measurer.style.cssText = 'position:fixed;left:-99999px;top:-99999px;visibility:hidden;white-space:nowrap;';
var probes = [];
fos.forEach(function(fo) {
  var inner = fo.querySelector('div');
  if (!inner) { probes.push(null); return; }
  var probe = document.createElement('div');
  probe.style.cssText = 'display:inline-block;white-space:nowrap;';
  probe.innerHTML = inner.textContent;  // 改用 textContent（同時修 SEC-001）
  measurer.appendChild(probe);
  probes.push({fo: fo, inner: inner, probe: probe});
});
document.body.appendChild(measurer);
// 現在一次性讀所有 offsetWidth（單一 reflow）
probes.forEach(function(p) {
  if (!p) return;
  var w = Math.ceil(p.probe.offsetWidth || p.probe.scrollWidth || 0);
  // ...apply
});
document.body.removeChild(measurer);
```

---

### PERF-003 [Medium] `_liftLabels` 呼叫鏈 — 每次 render 固定觸發：`_liftLabels` → `_growClusterLabels` → `_detectOverlaps` (via rAF)

**位置:** `src/WebView2Manager.cpp`, `_processPendingMermaid` 與 `renderContent` 末段

```javascript
_liftEdgeLabels(container);                           // 呼叫 _liftLabels
container.querySelectorAll('.mermaid-container').forEach(_refreshAutoFixBtn); // rAF → _detectOverlaps
```

每次 `renderContent` 都做完整的 lift + grow + detect 三步，即使 SVG 內容沒有改變（例如主題切換後的 re-render）。

**建議優化:**
- `_growClusterLabels` 結果應 cache（per container id）。如果 SVG 的 `innerHTML` hash 沒變，跳過 grow。
- Theme switch（`switchTheme`）觸發的 re-render 只需重做 mermaid render，不需重新 grow/detect。

---

### PERF-004 [Medium] `UpdatePreview` — `m_renderDirty` 只在 Bun 路徑設置，WebView2-only 路徑無去抖動改善

**位置:** `src/MermaidPreview.cpp`, `UpdatePreview`, 第 761–765 行

```cpp
if (m_renderFuture.valid() &&
    m_renderFuture.wait_for(std::chrono::seconds(0)) != std::future_status::ready) {
    m_renderDirty = true;
    return;
}
```

`m_renderDirty = true` 時 IDT_BUN_POLL 會在 40ms 後重試 `OnBunRenderComplete`，進而呼叫 `UpdatePreview(m_hWndLastView)`。但此時 `m_hWndLastView` 可能已與使用者的實際焦點窗口不同（多 tab 切換）。`OnBunRenderComplete` 的 `if (m_hWndLastView && IsWindow(m_hWndLastView)) UpdatePreview(m_hWndLastView)` 沒有重新檢查是否為目前活躍的 tab。

**影響:** 若使用者在 Bun render 期間切換 tab，dirty re-render 會以舊 tab 的 view 觸發，造成短暫渲染錯 tab 的預覽（視覺閃爍或內容錯位）。

**建議:**
```cpp
// OnBunRenderComplete 重試前確認 view 仍有效且為 active
if (m_renderDirty) {
    m_renderDirty = false;
    HWND activeView = ...; // 重新取得當前 active view
    if (activeView && IsWindow(activeView))
        UpdatePreview(activeView);
}
```

---

### PERF-005 [Medium] `ExtractFlowchartNodes` — `alreadySeen` 使用線性 vector 搜尋

**位置:** `src/MarkdownParser.cpp`, `ExtractFlowchartNodes`, 第 132–136 行

```cpp
std::vector<std::wstring> seen;
auto alreadySeen = [&](const std::wstring& id) {
    for (const auto& s : seen) if (s == id) return true;
    return false;
};
```

**根因:** O(N²) 搜尋。對 100+ 節點的圖，`alreadySeen` 在內層迴圈被呼叫 N 次，每次最壞 N 次比對 = O(N²)。

**量化影響:** 100 節點圖 = 10,000 次字串比對。每個 `std::wstring` 比對 ~10 chars，估計 0.1ms —對實際渲染影響較小，但在 node-edit 路徑（每次 inline edit 都呼叫）可積累。

**建議修復:** 改用 `std::unordered_set<std::wstring>`：
```cpp
std::unordered_set<std::wstring> seen;
auto alreadySeen = [&](const std::wstring& id) {
    return seen.count(id) > 0;
};
// ...
seen.insert(nodeId);
```

---

### PERF-006 [Medium] `_detectOverlaps` `getBoundingClientRect` 結果未 cache，`_applyAutoFix` 呼叫第二次

**位置:** `src/WebView2Manager.cpp`, `_applyAutoFix`

```javascript
async function _applyAutoFix(container) {
  // ...
  var hits = _detectOverlaps(container);   // ← 呼叫一次（第二次！）
  // _refreshAutoFixBtnNow 已經呼叫過一次 _detectOverlaps
```

使用者點擊 Auto-fix 按鈕時，`_applyAutoFix` 再次呼叫 `_detectOverlaps`，重複執行整個 N×M AABB 計算。第一次是在 `_refreshAutoFixBtnNow` 中，第二次在 `_applyAutoFix` 中。

**建議:** `_refreshAutoFixBtnNow` 把 `hits` 存在 container 的 dataset 或閉包變數中，`_applyAutoFix` 讀取 cached hits。

---

## 3. 程式碼品質 (Code Quality)

### QUALITY-001 [Low] `_injectSpacingDirective` regex 對 `%%{init...}%%` 的 `[^%]*` 過寬

**位置:** `src/WebView2Manager.cpp`, `_injectSpacingDirective`

```javascript
var hasDirective = /%%\{[^%]*init[^%]*\}%%/.test(src);
```

`[^%]*` 允許 `{` `}` 在內，因此若 mermaid 文件含有多個 `%%{...}%%` 指令，正則可能跨越指令邊界匹配，導致誤判 hasDirective 為 true。應改為非貪婪或更嚴格的 pattern：
```javascript
var hasDirective = /%%\{[^%{}]*init[^%{}]*\}%%/.test(src);
```

---

### QUALITY-002 [Low] `extractStrField` 與 `extractStr` 重複實作（DRY 違反）

**位置:** `src/WebView2Manager.cpp`, `editMermaidBlock` 與 `editMermaidNode` 分支

兩個幾乎完全相同的 lambda，差別僅在名稱。可提升為一個有名稱的私有成員函式或 module-level static。

---

### QUALITY-003 [Low] `_liftLabels` 呼叫 `_growClusterLabels`，`_liftEdgeLabels` 是別名 — 文件不一致

```javascript
var _liftEdgeLabels = _liftLabels;
```

外部有些地方呼叫 `_liftEdgeLabels(container)`，有些地方呼叫 `_liftLabels(root)` — 語意上等同，但命名混亂。`_liftEdgeLabels` 實際上也 lift cluster labels（透過 `_growClusterLabels`），名字具誤導性。建議統一為 `_liftLabels` 並移除別名。

---

## 4. 生命期管理補充確認

下列項目在此次審查中確認**無新問題**（延續 v2/v3 已修補狀態）：
- `_growClusterLabels` 的 measurer 在 `finally` 塊確保移除 ✅（error path 已覆蓋）
- `m_renderFuture` 在 `~CMermaidFrame` 與 `CloseCustomBar` 均有 detach 到背景 thread ✅
- `m_renderDirty` 在 `CloseCustomBar` 重置 ✅
- `SetEditMermaidNodeCallback` / `SetEditMermaidBlockCallback` callback 在 `WebView2Manager::Destroy` 路徑安全（`m_bDestroyed` flag 防止 use-after-free）✅

---

## 5. Actionable Task List

### Critical Path（立即修復）
- [ ] **[SEC-001]** `_growClusterLabels`: `probe.innerHTML = inner.innerHTML` → 改為 `probe.textContent = inner.textContent`（同時修 PERF-002 批次化）(`WebView2Manager.cpp`, `_growClusterLabels`)
- [ ] **[SEC-002]** `_wrapLongSubgraphTitles`: 加 title sanitization，拒絕含 `%%`/`-->`/newline 的 title；驗證插入的 `<br/>` 是否為 mermaid 合法換行 (`WebView2Manager.cpp`, `_wrapLongSubgraphTitles`)

### High Priority（Release 前修復）
- [ ] **[SEC-003]** `editMermaidNode` nodeId 加長度上限（128 chars）(`WebView2Manager.cpp`, 第 1468 行附近)
- [ ] **[SEC-004]** 手寫 JSON parser 補 `\uXXXX` 解碼（兩個 extractStr lambda 都需要更新）(`WebView2Manager.cpp`)
- [ ] **[PERF-001]** `_detectOverlaps` 批次讀取所有 `getBoundingClientRect` 以減少強制 reflow 次數
- [ ] **[PERF-002]** `_growClusterLabels` 批次化測量（所有 probe 一次 appendChild，一次 offsetWidth 讀取，一次 removeChild）

### Performance Tasks
- [ ] **[PERF-003]** `_growClusterLabels` 加 per-container SVG hash cache，主題切換不重跑 grow (`EstimatedImpact: -30% re-render overhead`)
- [ ] **[PERF-004]** `OnBunRenderComplete` dirty re-render 重新取得 active view 再呼叫 `UpdatePreview`
- [ ] **[PERF-005]** `ExtractFlowchartNodes` `alreadySeen` 改用 `std::unordered_set`
- [ ] **[PERF-006]** `_applyAutoFix` 讀取 cached hits 而非重新呼叫 `_detectOverlaps`

### Quality Tasks
- [ ] **[QUALITY-001]** `_injectSpacingDirective` regex 改為 `[^%{}]*` 防止跨 directive 誤匹配
- [ ] **[QUALITY-002]** 合併 `extractStrField` 與 `extractStr` 為單一函式
- [ ] **[QUALITY-003]** 移除 `_liftEdgeLabels` 別名，統一使用 `_liftLabels`

---

## 6. 本輪新問題總結

本輪共發現 **14 個新問題**（不含 v2/v3 已修補項）：
- Critical: 2（SEC-001, SEC-002）
- High: 4（SEC-003, SEC-004, PERF-001, PERF-002）
- Medium: 5（SEC-005, SEC-006, PERF-003, PERF-004, PERF-005）
- Low: 3（QUALITY-001, QUALITY-002, QUALITY-003）
