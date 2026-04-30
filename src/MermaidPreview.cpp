#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>

#ifndef VERIFY
#ifdef _DEBUG
#define VERIFY(x) _ASSERT(x)
#else
#define VERIFY(x) (void)(x)
#endif
#endif

#define EE_EXTERN_ONLY
#define ETL_FRAME_CLASS_NAME CMermaidFrame
#include "etlframe.h"

#include "MermaidPreview.h"
#include "WebView2Manager.h"
#include "BunRenderer.h"
#include "MarkdownParser.h"
#include "resource.h"
#include <functional>
#include <chrono>
#include <thread>

// ============================================================================
// Custom Bar host window class name
// ============================================================================
static const wchar_t* const kHostClassName = L"MermaidPreviewHost";
static bool s_bHostClassRegistered = false;

static const wchar_t* const kParkingClassName = L"MermaidPreviewParking";
static bool s_bParkingClassRegistered = false;

static CMermaidFrame* GetFrameFromHost(HWND hwnd)
{
    return reinterpret_cast<CMermaidFrame*>(GetWindowLongPtr(hwnd, GWLP_USERDATA));
}

// ============================================================================
// Host Window Procedure
// ============================================================================
LRESULT CALLBACK CMermaidFrame::HostWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg) {
    case WM_SIZE: {
        CMermaidFrame* pFrame = GetFrameFromHost(hwnd);
        if (pFrame && pFrame->m_pWebView) {
            RECT rc;
            GetClientRect(hwnd, &rc);
            pFrame->m_pWebView->Resize(rc);
        }
        return 0;
    }
    case WM_TIMER: {
        if (wParam == IDT_DEBOUNCE) {
            KillTimer(hwnd, IDT_DEBOUNCE);
            CMermaidFrame* pFrame = GetFrameFromHost(hwnd);
            if (pFrame && pFrame->m_hWndLastView &&
                IsWindow(pFrame->m_hWndLastView)) {
                pFrame->UpdatePreview(pFrame->m_hWndLastView);
            }
            return 0;
        }
        if (wParam == IDT_SCROLL_SYNC) {
            KillTimer(hwnd, IDT_SCROLL_SYNC);
            CMermaidFrame* pFrame = GetFrameFromHost(hwnd);
            if (pFrame && pFrame->m_hWndLastView &&
                IsWindow(pFrame->m_hWndLastView)) {
                pFrame->SyncScrollToPreview(pFrame->m_hWndLastView);
            }
            return 0;
        }
        if (wParam == IDT_SYNC_RESET_E2P) {
            KillTimer(hwnd, IDT_SYNC_RESET_E2P);
            CMermaidFrame* pFrame = GetFrameFromHost(hwnd);
            if (pFrame) pFrame->m_bSyncFromEditor = false;
            return 0;
        }
        if (wParam == IDT_SYNC_RESET_P2E) {
            KillTimer(hwnd, IDT_SYNC_RESET_P2E);
            CMermaidFrame* pFrame = GetFrameFromHost(hwnd);
            if (pFrame) pFrame->m_bSyncFromPreview = false;
            return 0;
        }
        if (wParam == IDT_BUN_POLL) {
            CMermaidFrame* pFrame = GetFrameFromHost(hwnd);
            if (pFrame) pFrame->OnBunRenderComplete();
            return 0;
        }
        break;
    }
    case WM_DESTROY:
        return 0;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

// ============================================================================
// ~CMermaidFrame - Detach any in-flight Bun render future before it would
// otherwise block this destructor for up to 15 s. The worker holds a
// shared_ptr<BunRenderer> so the renderer survives until the pipe drains.
// ============================================================================
CMermaidFrame::~CMermaidFrame()
{
    if (m_renderFuture.valid()) {
        std::thread([f = std::move(m_renderFuture)]() mutable {
            try { f.wait(); } catch (...) {}
        }).detach();
    }
}

// ============================================================================
// OnCommand - Toggle the sidebar
// ============================================================================
void CMermaidFrame::OnCommand(HWND hwndView)
{
    if (m_bVisible) {
        CloseCustomBar(hwndView);
        m_bAutoOpened = false;
    } else {
        OpenCustomBar(hwndView);
        m_bAutoOpened = false;
    }
}

// ============================================================================
// QueryStatus
// ============================================================================
BOOL CMermaidFrame::QueryStatus(HWND /*hwndView*/, LPBOOL pbChecked)
{
    *pbChecked = m_bVisible ? TRUE : FALSE;
    return TRUE;
}

// ============================================================================
// OnEvents
// ============================================================================
void CMermaidFrame::OnEvents(HWND hwndView, UINT nEvent, LPARAM lParam)
{
    if (nEvent & EVENT_CREATE_FRAME) {
        LoadSettings();
        // Create a 1x1 hidden popup to serve as parking window for WebView2
        if (!s_bParkingClassRegistered) {
            WNDCLASSEX wc = {};
            wc.cbSize = sizeof(wc);
            wc.lpfnWndProc = DefWindowProc;
            wc.hInstance = EEGetInstanceHandle();
            wc.lpszClassName = kParkingClassName;
            if (RegisterClassEx(&wc))
                s_bParkingClassRegistered = true;
        }
        m_hwndParking = CreateWindowEx(0, kParkingClassName,
            L"MermaidPreviewParking", WS_POPUP,
            0, 0, 1, 1,
            nullptr, nullptr, EEGetInstanceHandle(), nullptr);
        return;
    }

    if (nEvent & EVENT_CLOSE_FRAME) {
        SaveSettings();
        // Full cleanup: destroy WebView2 regardless of visible/parked state
        if (m_pWebView) {
            m_pWebView->Destroy();
            m_pWebView.reset();
        }
        if (m_hwndHost) {
            DestroyWindow(m_hwndHost);
            m_hwndHost = nullptr;
        }
        if (m_nBarID != 0) {
            Editor_CustomBarClose(m_hWnd, m_nBarID);
            m_nBarID = 0;
        }
        // Destroy parking window last
        if (m_hwndParking) {
            DestroyWindow(m_hwndParking);
            m_hwndParking = nullptr;
        }
        m_bVisible = false;
        m_bParked = false;
        m_bAutoOpened = false;
        // Wait for async Bun startup (with timeout to prevent hanging EmEditor)
        if (m_bunStartFuture.valid()) {
            auto status = m_bunStartFuture.wait_for(std::chrono::milliseconds(2000));
            if (status == std::future_status::ready) {
                m_bunStartFuture.get(); // consume the future
            }
            // If timeout: Bun is stuck — proceed with cleanup anyway.
            // BunRenderer::Stop() will terminate the process.
        }
        if (m_pBunRenderer) {
            m_pBunRenderer->Stop();
            m_pBunRenderer.reset();
        }
        m_bBunAvailable = false;
        return;
    }

    if (nEvent & EVENT_CUSTOM_BAR_CLOSED) {
        OnCustomBarClosed(hwndView, lParam);
        return;
    }

    if (nEvent & EVENT_FILE_OPENED) {
        m_hWndLastView = hwndView;
        if (!m_bVisible) {
            TryAutoOpen(hwndView);
        } else {
            m_nLastHash = 0;
            m_sLastContent.clear();
            UpdatePreview(hwndView);
        }
        return;
    }

    // EVENT_DOC_CLOSE: close (park) preview immediately.
    // No return — fall through so EVENT_DOC_SEL_CHANGED can also run,
    // but the bitmask guard below prevents the costly close→reopen cycle.
    if (nEvent & EVENT_DOC_CLOSE) {
        // Clear dangling HWND: this view is about to be destroyed by EmEditor.
        if (hwndView == m_hWndLastView)
            m_hWndLastView = nullptr;
        if (m_bVisible) {
            CloseCustomBar(hwndView);
        }
    }

    if (nEvent & EVENT_DOC_SEL_CHANGED) {
        m_hWndLastView = hwndView;
        if (!m_bVisible) {
            // Guard: if EVENT_DOC_CLOSE is in the same bitmask, skip auto-open
            // to avoid the expensive close→reopen cycle that blocks EmEditor.
            if (!(nEvent & EVENT_DOC_CLOSE)) {
                TryAutoOpen(hwndView);
            }
        } else {
            // Preview is visible — check if new tab still has mermaid
            if (!IsMarkdownFile(hwndView) || !HasMermaidBlocks(hwndView)) {
                CloseCustomBar(hwndView);
            } else {
                m_nLastHash = 0;
                m_sLastContent.clear();
                UpdatePreview(hwndView);
            }
        }
        return;
    }

    if (!m_bVisible)
        return;

    m_hWndLastView = hwndView;

    // NOTE: nEvent is a bitmask — multiple events can fire simultaneously
    // (e.g. EVENT_SCROLL | EVENT_MODIFIED when typing causes scroll).
    // Do NOT use early return here; process each flag independently.

    if (nEvent & EVENT_SCROLL) {
        if (!m_bSyncFromPreview && m_hwndHost) {
            KillTimer(m_hwndHost, IDT_SCROLL_SYNC);
            SetTimer(m_hwndHost, IDT_SCROLL_SYNC, SCROLL_SYNC_MS, nullptr);
        }
    }

    if (nEvent & EVENT_MODIFIED) {
        if (m_hwndHost) {
            KillTimer(m_hwndHost, IDT_DEBOUNCE);
            SetTimer(m_hwndHost, IDT_DEBOUNCE, DEBOUNCE_MS, nullptr);
        }
    }

    if (nEvent & EVENT_UI_CHANGED) {
        if (!m_bDarkModeOverride) {
            bool dark = IsDarkMode(hwndView);
            if (dark != m_bDarkMode) {
                m_bDarkMode = dark;
                if (m_pWebView && m_pWebView->IsReady()) {
                    m_pWebView->SetTheme(m_bDarkMode);
                    m_nLastHash = 0;
                    m_sLastContent.clear();
                    UpdatePreview(hwndView);
                }
            }
        }
    }
}

// ============================================================================
// EnsureBunRenderer - Lazily start Bun in a background thread (non-blocking)
// ============================================================================
void CMermaidFrame::EnsureBunRenderer()
{
    if (m_bBunAvailable)
        return;

    // Check if an async start has completed
    if (m_bunStartFuture.valid()) {
        if (m_bunStartFuture.wait_for(std::chrono::seconds(0)) == std::future_status::ready) {
            m_bBunAvailable = m_bunStartFuture.get();
        }
        return; // Still starting or just finished
    }

    if (!m_pBunRenderer) {
        m_pBunRenderer = std::make_shared<BunRenderer>();
    }

    // Launch Bun startup in background thread to avoid freezing UI
    // Capture shared_ptr to prevent dangling pointer if m_pBunRenderer is reset
    auto rendererPtr = m_pBunRenderer;
    m_bunStartFuture = std::async(std::launch::async, [rendererPtr]() {
        return rendererPtr->Start();
    });
}

// ============================================================================
// OpenCustomBar
// ============================================================================
void CMermaidFrame::OpenCustomBar(HWND hwndView, std::wstring prefetchedContent)
{
    if (m_bVisible)
        return;

    if (!s_bHostClassRegistered) {
        WNDCLASSEX wc = {};
        wc.cbSize = sizeof(wc);
        wc.lpfnWndProc = HostWndProc;
        wc.hInstance = EEGetInstanceHandle();
        wc.lpszClassName = kHostClassName;
        wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
        wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
        if (RegisterClassEx(&wc))
            s_bHostClassRegistered = true;
    }

    m_hwndHost = CreateWindowEx(
        0,
        kHostClassName,
        L"Mermaid Preview",
        WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN,
        0, 0, 300, 400,
        m_hWnd,
        nullptr,
        EEGetInstanceHandle(),
        nullptr);

    if (!m_hwndHost)
        return;

    SetWindowLongPtr(m_hwndHost, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(this));

    CUSTOM_BAR_INFO cbi = {};
    cbi.cbSize = sizeof(cbi);
    cbi.hwndClient = m_hwndHost;
    cbi.pszTitle = L"Mermaid Preview";
    cbi.iPos = m_iBarPos;

    m_nBarID = Editor_CustomBarOpen(m_hWnd, &cbi);
    if (m_nBarID == 0) {
        DestroyWindow(m_hwndHost);
        m_hwndHost = nullptr;
        return;
    }

    m_bVisible = true;
    m_hWndLastView = hwndView;
    if (!m_bDarkModeOverride) {
        m_bDarkMode = IsDarkMode(hwndView);
    }

    // Try to start Bun renderer in background
    EnsureBunRenderer();

    // === FAST PATH: Reparent parked WebView2 (~10ms vs ~800-1500ms) ===
    if (m_bParked && m_pWebView && m_pWebView->IsReady()) {
        HRESULT hr = m_pWebView->Reparent(m_hwndHost);
        if (SUCCEEDED(hr)) {
            m_bParked = false;
            m_pWebView->SetTheme(m_bDarkMode);
            // Force re-render with current content
            m_nLastHash = 0;
            m_sLastContent.clear();
            UpdatePreview(hwndView);
            return;
        }
        // Reparent failed — fall through to full init
        m_pWebView->Destroy();
        m_pWebView.reset();
        m_bParked = false;
    }

    // === FALLBACK: parked but not ready (init was in progress) ===
    if (m_bParked && m_pWebView) {
        m_pWebView->Destroy();
        m_pWebView.reset();
        m_bParked = false;
    }

    // === FULL INIT: Create new WebView2 ===
    m_pWebView = std::make_unique<WebView2Manager>();
    m_pWebView->Initialize(m_hwndHost, [this]() {
        if (m_pWebView) {
            m_pWebView->SetTheme(m_bDarkMode);

            // Register WebMessage handler for editing (uses m_hWndLastView, not captured hwndView)
            m_pWebView->SetEditCallback([this](int lineStart, int lineEnd, const std::wstring& newText) {
                if (m_hWndLastView && IsWindow(m_hWndLastView))
                    OnPreviewTextEdited(m_hWndLastView, lineStart, lineEnd, newText);
            });

            // M2: register mermaid-node label inline-edit handler.
            m_pWebView->SetEditMermaidNodeCallback(
                [this](const std::wstring& blockId,
                       const std::wstring& nodeId,
                       const std::wstring& newLabel) {
                    if (m_hWndLastView && IsWindow(m_hWndLastView))
                        OnPreviewMermaidNodeEdited(m_hWndLastView, blockId, nodeId, newLabel);
                });

            // Phase 1: register whole-block replacement (auto-correction).
            m_pWebView->SetEditMermaidBlockCallback(
                [this](const std::wstring& blockId, const std::wstring& newSource) {
                    if (m_hWndLastView && IsWindow(m_hWndLastView))
                        OnPreviewMermaidBlockEdited(m_hWndLastView, blockId, newSource);
                });

            // Register theme change callback (from right-click context menu)
            m_pWebView->SetThemeCallback([this](bool dark) {
                m_bDarkMode = dark;
                m_bDarkModeOverride = true;
                // Force re-render on next update (Bun theme sync)
                m_nLastHash = 0;
                m_sLastContent.clear();
            });

            // Register scroll sync callback (Preview → Editor)
            m_pWebView->SetScrollCallback([this](int line) {
                if (m_hWndLastView && IsWindow(m_hWndLastView)) {
                    OnPreviewScrolled(m_hWndLastView, line);
                }
            });

            // Register navigate callback (click mermaid → jump editor to source)
            m_pWebView->SetNavigateCallback([this](int line) {
                if (m_hWndLastView && IsWindow(m_hWndLastView)) {
                    OnPreviewNavigate(m_hWndLastView, line);
                }
            });

            // Register open-file callback (relative path link clicked)
            m_pWebView->SetOpenFileCallback([this](const std::wstring& path) {
                if (m_hWndLastView && IsWindow(m_hWndLastView))
                    OnOpenFileLink(m_hWndLastView, path);
            });

            // Register font size callback (from right-click context menu)
            m_pWebView->SetFontSizeCallback([this](int size) {
                m_iFontSize = size;
                SaveSettings();
            });

            // Restore font size if non-default
            if (m_iFontSize != 14) {
                m_pWebView->SetFontSize(m_iFontSize);
            }

            // Optimization 3: Use pre-fetched content if available
            if (m_bHasPrefetch) {
                m_pWebView->RenderContent(m_sPrefetchedHtml, m_bDarkMode);
                if (m_hWndLastView && IsWindow(m_hWndLastView))
                    SyncScrollToPreview(m_hWndLastView);
                m_bHasPrefetch = false;
                m_sPrefetchedHtml.clear();
            } else if (m_hWndLastView && IsWindow(m_hWndLastView)) {
                UpdatePreview(m_hWndLastView);
            }
        }
    });

    // Optimization 3: Pre-fetch document content while WebView2 initializes async
    // This overlaps content preparation with the ~800-1500ms WebView2 startup
    {
        // Use prefetched content if provided (from TryAutoOpen), else read fresh
        std::wstring content = prefetchedContent.empty()
            ? MarkdownParser::GetDocumentContent(hwndView)
            : std::move(prefetchedContent);
        std::hash<std::wstring> hasher;
        m_nLastHash = hasher(content);
        m_sLastContent = content;
        m_sPrefetchedHtml = MarkdownParser::ConvertToHtml(content);
        m_bHasPrefetch = true;
    }
}

// ============================================================================
// CloseCustomBar
// ============================================================================
void CMermaidFrame::CloseCustomBar(HWND /*hwndView*/)
{
    if (!m_bVisible)
        return;

    // 1. Kill timers
    if (m_hwndHost) {
        KillTimer(m_hwndHost, IDT_DEBOUNCE);
        KillTimer(m_hwndHost, IDT_SCROLL_SYNC);
        KillTimer(m_hwndHost, IDT_SYNC_RESET_E2P);
        KillTimer(m_hwndHost, IDT_SYNC_RESET_P2E);
        KillTimer(m_hwndHost, IDT_BUN_POLL);
    }
    m_renderDirty = false;
    m_renderPendingHtml.clear();
    m_renderPendingView = nullptr;

    // Detach any in-flight Bun render. std::async(launch::async) futures
    // *block in their destructor* until the shared state is ready — moving
    // the future into a self-joining thread keeps the worker running but
    // unblocks the UI close path. The worker holds a shared_ptr<BunRenderer>,
    // so the renderer stays alive until the pipe drains naturally.
    if (m_renderFuture.valid()) {
        std::thread([f = std::move(m_renderFuture)]() mutable {
            try { f.wait(); } catch (...) {}
        }).detach();
    }

    // 2. Restore focus to EmEditor BEFORE parking WebView2.
    //    WebView2 browser process may own the focus; reclaim it first
    //    so that parking doesn't leave focus on the hidden parking window.
    if (m_hWnd && IsWindow(m_hWnd))
        SetFocus(m_hWnd);

    // 3. Park WebView2 to parking window (if controller exists)
    if (m_pWebView && m_pWebView->HasController() && m_hwndParking) {
        m_pWebView->Park(m_hwndParking);
        m_bParked = true;
    } else if (m_pWebView) {
        // Controller not yet created (async init in progress) — destroy
        m_pWebView->Destroy();
        m_pWebView.reset();
        m_bParked = false;
    }

    // 4. Close custom bar
    if (m_nBarID != 0) {
        Editor_CustomBarClose(m_hWnd, m_nBarID);
        m_nBarID = 0;
    }

    // 5. Destroy host window (WebView2 is already parked elsewhere, safe)
    if (m_hwndHost) {
        DestroyWindow(m_hwndHost);
        m_hwndHost = nullptr;
    }

    m_bVisible = false;
    m_bSyncFromEditor = false;
    m_bSyncFromPreview = false;
    m_sLastContent.clear();
    m_nLastHash = 0;
}

// ============================================================================
// OnCustomBarClosed
// ============================================================================
void CMermaidFrame::OnCustomBarClosed(HWND /*hwndView*/, LPARAM lParam)
{
    auto* pInfo = reinterpret_cast<CUSTOM_BAR_CLOSE_INFO*>(lParam);
    if (!pInfo || pInfo->nID != m_nBarID)
        return;

    m_iBarPos = pInfo->iPos;

    if (m_hwndHost) {
        KillTimer(m_hwndHost, IDT_DEBOUNCE);
        KillTimer(m_hwndHost, IDT_SCROLL_SYNC);
        KillTimer(m_hwndHost, IDT_SYNC_RESET_E2P);
        KillTimer(m_hwndHost, IDT_SYNC_RESET_P2E);
        KillTimer(m_hwndHost, IDT_BUN_POLL);
    }
    m_renderDirty = false;
    m_renderPendingHtml.clear();
    m_renderPendingView = nullptr;

    // Restore focus before parking (WebView2 browser process may own focus)
    if (m_hWnd && IsWindow(m_hWnd))
        SetFocus(m_hWnd);

    // Park WebView2 (guard: !m_bParked prevents double-park if CloseCustomBar already did it)
    if (!m_bParked && m_pWebView && m_pWebView->HasController() && m_hwndParking) {
        m_pWebView->Park(m_hwndParking);
        m_bParked = true;
    } else if (!m_bParked && m_pWebView) {
        m_pWebView->Destroy();
        m_pWebView.reset();
    }

    if (m_hwndHost) {
        DestroyWindow(m_hwndHost);
        m_hwndHost = nullptr;
    }

    m_nBarID = 0;
    m_bVisible = false;
    m_bSyncFromEditor = false;
    m_bSyncFromPreview = false;
    m_sLastContent.clear();
    m_nLastHash = 0;

    SaveSettings();
}

// ============================================================================
// IsMarkdownFile
// ============================================================================
bool CMermaidFrame::IsMarkdownFile(HWND hwndView) const
{
    if (!hwndView || !IsWindow(hwndView))
        return false;

    WCHAR szPath[MAX_PATH] = {};
    Editor_Info(hwndView, EI_GET_FILE_NAMEW, (LPARAM)szPath);

    if (szPath[0] == L'\0')
        return false;

    LPCWSTR pExt = wcsrchr(szPath, L'.');
    if (!pExt)
        return false;

    return (_wcsicmp(pExt, L".md") == 0 ||
            _wcsicmp(pExt, L".markdown") == 0 ||
            _wcsicmp(pExt, L".mmd") == 0 ||
            _wcsicmp(pExt, L".mermaid") == 0);
}

// ============================================================================
// HasMermaidBlocks
// ============================================================================
bool CMermaidFrame::HasMermaidBlocks(HWND hwndView) const
{
    if (!hwndView || !IsWindow(hwndView))
        return false;

    std::wstring content = MarkdownParser::GetDocumentContent(hwndView);
    auto blocks = MarkdownParser::ExtractMermaidBlocks(content);
    return !blocks.empty();
}

// TryAutoClose was wired to no caller (audit v2 LOW-4.4). The auto-close
// behaviour is undesirable in practice — a user transiently deleting a
// mermaid block while editing should not yank the preview panel — so the
// dead function is intentionally removed. Manual close still works via
// the toolbar button.

void CMermaidFrame::TryAutoOpen(HWND hwndView)
{
    if (m_bVisible) return;
    if (!IsMarkdownFile(hwndView)) return;

    // Read content once (avoids double read in HasMermaidBlocks + OpenCustomBar prefetch)
    std::wstring content = MarkdownParser::GetDocumentContent(hwndView);
    auto blocks = MarkdownParser::ExtractMermaidBlocks(content);
    if (!blocks.empty()) {
        OpenCustomBar(hwndView, std::move(content));
        m_bAutoOpened = true;
    }
}

// ============================================================================
// SpliceSvgIntoHtml - Replace `<div class="mermaid-container">` placeholders
// with the SVG (or error block) that Bun produced. Pure string surgery; safe
// to call on either the UI thread or after a background render completes.
// ============================================================================
void CMermaidFrame::SpliceSvgIntoHtml(std::wstring& html,
                                      const std::vector<MermaidRenderResult>& results)
{
    for (auto& r : results) {
        std::wstring placeholder = L"data-mermaid-id=\"" + r.id + L"\"";
        size_t pos = html.find(placeholder);
        if (pos == std::wstring::npos) continue;

        size_t divStart = html.rfind(L"<div ", pos);
        size_t divEnd = html.find(L"</div>", pos);
        if (divStart == std::wstring::npos || divEnd == std::wstring::npos) continue;
        divEnd += 6;

        std::wstring origDiv = html.substr(divStart, divEnd - divStart);
        std::wstring dataSrc, dataLineStart, dataLineEnd;
        size_t srcAttr = origDiv.find(L"data-mermaid-src=\"");
        if (srcAttr != std::wstring::npos) {
            srcAttr += 18;
            size_t srcEnd = origDiv.find(L'"', srcAttr);
            if (srcEnd != std::wstring::npos)
                dataSrc = origDiv.substr(srcAttr, srcEnd - srcAttr);
        }
        size_t lsAttr = origDiv.find(L"data-line-start=\"");
        if (lsAttr != std::wstring::npos) {
            lsAttr += 17;
            size_t lsEnd = origDiv.find(L'"', lsAttr);
            if (lsEnd != std::wstring::npos)
                dataLineStart = origDiv.substr(lsAttr, lsEnd - lsAttr);
        }
        size_t leAttr = origDiv.find(L"data-line-end=\"");
        if (leAttr != std::wstring::npos) {
            leAttr += 15;
            size_t leEnd = origDiv.find(L'"', leAttr);
            if (leEnd != std::wstring::npos)
                dataLineEnd = origDiv.substr(leAttr, leEnd - leAttr);
        }

        std::wstring attrs = L"class=\"mermaid-container\" data-mermaid-id=\"" + r.id + L"\"";
        if (!dataSrc.empty())       attrs += L" data-mermaid-src=\"" + dataSrc + L"\"";
        if (!dataLineStart.empty()) attrs += L" data-line-start=\"" + dataLineStart + L"\"";
        if (!dataLineEnd.empty())   attrs += L" data-line-end=\"" + dataLineEnd + L"\"";

        if (!r.svg.empty()) {
            std::wstring svgDiv = L"<div " + attrs + L">" + r.svg + L"</div>";
            html = html.substr(0, divStart) + svgDiv + html.substr(divEnd);
        } else if (!r.error.empty()) {
            std::wstring errDiv = L"<div " + attrs + L"><div class=\"mermaid-error\">Mermaid error: "
                + MarkdownParser::HtmlEscape(r.error) + L"</div></div>";
            html = html.substr(0, divStart) + errDiv + html.substr(divEnd);
        }
    }
}

// ============================================================================
// UpdatePreview - Hybrid: C++ markdown + Bun mermaid SVG (fallback: WebView2 JS)
//
// Bun's RenderBlocks call can take ~50–500 ms (or up to 15 s if Bun hangs),
// so it must NOT run on the UI thread. The flow:
//
//   1. Parse markdown → HTML synchronously (cheap, ~1 ms).
//   2. If a previous Bun job is still running, set m_renderDirty and bail —
//      the polling timer (IDT_BUN_POLL) will re-trigger UpdatePreview after
//      the in-flight job completes, picking up the latest editor content.
//   3. Otherwise, render the placeholder HTML immediately so the user sees
//      text without waiting for Bun, then kick off RenderBlocks on a worker
//      thread via std::async + start IDT_BUN_POLL to splice in the SVG.
//   4. If Bun isn't available at all, render once and we're done (the JS-side
//      mermaid.js will handle the placeholders client-side).
// ============================================================================
void CMermaidFrame::UpdatePreview(HWND hwndView)
{
    if (!m_bVisible || !m_pWebView)
        return;

    if (!hwndView || !IsWindow(hwndView))
        return;

    // Check if async Bun startup has completed
    if (!m_bBunAvailable && m_bunStartFuture.valid()) {
        if (m_bunStartFuture.wait_for(std::chrono::seconds(0)) == std::future_status::ready) {
            m_bBunAvailable = m_bunStartFuture.get();
        }
    }

    // If a Bun render is already in flight, mark dirty and bail. The poll
    // timer will re-enter UpdatePreview when the future resolves.
    if (m_renderFuture.valid() &&
        m_renderFuture.wait_for(std::chrono::seconds(0)) != std::future_status::ready) {
        m_renderDirty = true;
        return;
    }

    std::wstring content = MarkdownParser::GetDocumentContent(hwndView);

    // Quick hash comparison
    std::hash<std::wstring> hasher;
    size_t h = hasher(content);
    if (h == m_nLastHash && content == m_sLastContent)
        return;

    m_nLastHash = h;
    m_sLastContent = content;

    // C++ native: convert markdown to HTML with line tracking
    std::wstring html = MarkdownParser::ConvertToHtml(content);

    // Decide whether to dispatch Bun
    bool useBun = m_bBunAvailable && m_pBunRenderer && m_pBunRenderer->IsReady();
    std::vector<MermaidBlock> mermaidBlocks;
    if (useBun) {
        mermaidBlocks = MarkdownParser::ExtractMermaidBlocks(content);
        if (mermaidBlocks.empty()) useBun = false;
    }

    if (!useBun) {
        // No Bun → ship HTML now; client-side mermaid.js handles placeholders.
        m_pWebView->RenderContent(html, m_bDarkMode);
        SyncScrollToPreview(hwndView);
        return;
    }

    // Show text + placeholders immediately (sub-second perceived latency).
    // The client-side mermaid.js will start rendering them; we'll overwrite
    // with server-side SVG when Bun completes.
    m_pWebView->RenderContent(html, m_bDarkMode);
    SyncScrollToPreview(hwndView);

    // Capture render context for the completion handler.
    m_renderPendingHtml = std::move(html);
    m_renderPendingDark = m_bDarkMode;
    m_renderPendingView = hwndView;

    std::vector<std::pair<std::wstring, std::wstring>> bunBlocks;
    bunBlocks.reserve(mermaidBlocks.size());
    for (size_t i = 0; i < mermaidBlocks.size(); i++) {
        bunBlocks.push_back({
            L"mermaid-placeholder-" + std::to_wstring(i),
            mermaidBlocks[i].code
        });
    }
    std::wstring theme = m_bDarkMode ? L"dark" : L"default";

    // Capture renderer by shared_ptr — keeps BunRenderer alive even if
    // CMermaidFrame is being torn down while the worker is mid-pipe.
    auto renderer = m_pBunRenderer;
    m_renderFuture = std::async(std::launch::async,
        [renderer, blocks = std::move(bunBlocks), theme]() {
            return renderer->RenderBlocks(blocks, theme);
        });

    if (m_hwndHost) {
        SetTimer(m_hwndHost, IDT_BUN_POLL, BUN_POLL_MS, nullptr);
    }
}

// ============================================================================
// OnBunRenderComplete - IDT_BUN_POLL fires every BUN_POLL_MS ms while a
// background Bun render is in flight. When the future resolves, splice the
// SVGs into the cached HTML and re-render. If the document changed while Bun
// was working (m_renderDirty), kick off another UpdatePreview pass.
// ============================================================================
void CMermaidFrame::OnBunRenderComplete()
{
    if (!m_renderFuture.valid()) {
        if (m_hwndHost) KillTimer(m_hwndHost, IDT_BUN_POLL);
        return;
    }
    if (m_renderFuture.wait_for(std::chrono::seconds(0)) != std::future_status::ready)
        return; // still working — keep polling

    std::vector<MermaidRenderResult> results;
    try {
        results = m_renderFuture.get();
    } catch (...) {
        // Swallow worker exceptions; client-side mermaid.js already drew
        // something on the placeholder path so the user isn't stuck.
        results.clear();
    }
    if (m_hwndHost) KillTimer(m_hwndHost, IDT_BUN_POLL);

    // Bun returned (or timed out). If we got SVGs, splice and re-render.
    if (!results.empty() && m_pWebView) {
        std::wstring html = std::move(m_renderPendingHtml);
        SpliceSvgIntoHtml(html, results);
        m_pWebView->RenderContent(html, m_renderPendingDark);
    }
    m_renderPendingHtml.clear();
    m_renderPendingView = nullptr;

    // If the editor changed while Bun was working, run another pass now
    // so the preview catches up with the latest content.
    if (m_renderDirty) {
        m_renderDirty = false;
        if (m_hWndLastView && IsWindow(m_hWndLastView))
            UpdatePreview(m_hWndLastView);
    }
}

// ============================================================================
// SyncScrollToPreview - Editor → Preview scroll sync
// ============================================================================
void CMermaidFrame::SyncScrollToPreview(HWND hwndView)
{
    if (!m_pWebView || !m_pWebView->IsReady())
        return;
    if (!hwndView || !IsWindow(hwndView))
        return;

    POINT_PTR pt = {};
    Editor_GetScrollPos(hwndView, &pt);
    int topLine = (int)pt.y;

    m_bSyncFromEditor = true;
    m_pWebView->ScrollToLine(topLine);

    // Reset flag after 150ms to re-enable Preview→Editor sync
    if (m_hwndHost) {
        KillTimer(m_hwndHost, IDT_SYNC_RESET_E2P);
        SetTimer(m_hwndHost, IDT_SYNC_RESET_E2P, SYNC_RESET_MS, nullptr);
    }
}

// ============================================================================
// OnPreviewScrolled - Preview → Editor scroll sync
// ============================================================================
void CMermaidFrame::OnPreviewScrolled(HWND hwndView, int line)
{
    if (m_bSyncFromEditor)
        return;
    if (!hwndView || !IsWindow(hwndView))
        return;

    m_bSyncFromPreview = true;
    POINT_PTR pt = {};
    pt.y = line;
    Editor_SetScrollPos(hwndView, &pt);

    // Reset flag after 150ms to re-enable Editor→Preview sync
    if (m_hwndHost) {
        KillTimer(m_hwndHost, IDT_SYNC_RESET_P2E);
        SetTimer(m_hwndHost, IDT_SYNC_RESET_P2E, SYNC_RESET_MS, nullptr);
    }
}

// ============================================================================
// OnPreviewNavigate - Click mermaid diagram → jump editor to source line
// ============================================================================
void CMermaidFrame::OnPreviewNavigate(HWND hwndView, int line)
{
    if (!hwndView || !IsWindow(hwndView))
        return;

    // Set caret to the beginning of the target line
    POINT_PTR pt = {};
    pt.x = 0;
    pt.y = line;
    Editor_SetCaretPos(hwndView, POS_LOGICAL_W, &pt);

    // Scroll to make the target line visible (center-ish)
    POINT_PTR scroll = {};
    int scrollLine = (line > 5) ? line - 5 : 0;
    scroll.y = scrollLine;
    Editor_SetScrollPos(hwndView, &scroll);

    // Suppress scroll sync feedback
    m_bSyncFromPreview = true;
    if (m_hwndHost) {
        KillTimer(m_hwndHost, IDT_SYNC_RESET_P2E);
        SetTimer(m_hwndHost, IDT_SYNC_RESET_P2E, SYNC_RESET_MS, nullptr);
    }
}

// ============================================================================
// OnOpenFileLink - Open a relative-path file link in EmEditor
// ============================================================================
void CMermaidFrame::OnOpenFileLink(HWND hwndView, const std::wstring& relativePath)
{
    if (!hwndView || !IsWindow(hwndView))
        return;

    if (relativePath.empty())
        return;

    // 1. Get the current file's full path
    WCHAR szPath[MAX_PATH] = {};
    Editor_Info(hwndView, EI_GET_FILE_NAMEW, (LPARAM)szPath);
    if (szPath[0] == L'\0')
        return;

    // 2. Get the directory of the current file
    std::wstring dir = szPath;
    size_t lastSlash = dir.find_last_of(L"\\/");
    if (lastSlash == std::wstring::npos)
        return;
    dir = dir.substr(0, lastSlash);

    // 3. Convert forward slashes in relativePath to backslashes
    std::wstring relPath = relativePath;
    for (auto& c : relPath) {
        if (c == L'/') c = L'\\';
    }

    // 4. Combine: dir + relativePath
    std::wstring fullPath = dir + L"\\" + relPath;

    // 5. Canonicalize (resolve . and ..)
    WCHAR canonical[MAX_PATH] = {};
    DWORD len = GetFullPathNameW(fullPath.c_str(), MAX_PATH, canonical, nullptr);
    if (len == 0 || len >= MAX_PATH)
        return;

    // 5b. Boundary check: canonical must be inside the current document's
    //     directory (or a subdirectory). Blocks `../../etc/passwd` style
    //     traversal even after GetFullPathNameW collapsed the dots.
    WCHAR canonDir[MAX_PATH] = {};
    DWORD dirLen = GetFullPathNameW(dir.c_str(), MAX_PATH, canonDir, nullptr);
    if (dirLen == 0 || dirLen >= MAX_PATH)
        return;
    std::wstring canonStr = canonical;
    std::wstring canonDirStr = canonDir;
    // Normalise trailing slash on the directory side so "C:\\foo" matches
    // "C:\\foo\\bar.md" but not "C:\\foobar\\evil.md".
    if (canonDirStr.empty() || canonDirStr.back() != L'\\')
        canonDirStr += L'\\';
    if (canonStr.size() < canonDirStr.size())
        return;
    // Case-insensitive prefix check (Windows paths are case-insensitive).
    if (_wcsnicmp(canonStr.c_str(), canonDirStr.c_str(), canonDirStr.size()) != 0)
        return;

    // 6. Verify file exists
    DWORD attrs = GetFileAttributesW(canonical);
    if (attrs == INVALID_FILE_ATTRIBUTES || (attrs & FILE_ATTRIBUTE_DIRECTORY))
        return;

    // 7. Open file in EmEditor (new tab, not replacing current file)
    LOAD_FILE_INFO_EX lfi = {};
    lfi.cbSize = sizeof(lfi);
    lfi.nFlags = LFI_ALLOW_NEW_WINDOW;
    Editor_LoadFileW(m_hWnd, &lfi, canonical);
}

// ============================================================================
// OnPreviewTextEdited - Handle text edits from the preview panel
// ============================================================================
void CMermaidFrame::OnPreviewTextEdited(HWND hwndView, int lineStart, int lineEnd,
                                         const std::wstring& newText)
{
    if (!hwndView || !IsWindow(hwndView))
        return;

    // Use EmEditor API to select the line range and replace text.
    // POINT_PTR: x = column (0-based), y = line (0-based) for POS_LOGICAL_W

    UINT_PTR totalLines = (UINT_PTR)SendMessage(
        hwndView, EE_GET_LINES, (WPARAM)0, 0);

    // Validate line range from WebView2 (untrusted input)
    if (totalLines == 0) return;
    if (lineStart < 0 || lineStart >= (int)totalLines) return;
    if (lineEnd < lineStart) return;
    if (lineEnd >= (int)totalLines)
        lineEnd = (int)totalLines - 1;

    // Limit newText size (prevent massive paste from crashing EmEditor)
    constexpr size_t MAX_EDIT_LENGTH = 1024 * 1024; // 1MB
    if (newText.size() > MAX_EDIT_LENGTH) return;

    // Set caret at start of lineStart (no selection)
    POINT_PTR ptStart = {};
    ptStart.x = 0;
    ptStart.y = lineStart;
    Editor_SetCaretPos(hwndView, POS_LOGICAL_W, &ptStart);

    // Extend selection to start of the line after lineEnd
    // (this selects all content in lines lineStart..lineEnd including newlines)
    POINT_PTR ptEnd = {};
    ptEnd.x = 0;
    if (lineEnd + 1 < (int)totalLines) {
        ptEnd.y = lineEnd + 1;
    } else {
        // Last line: select to end of line
        GET_LINE_INFO gli = {};
        gli.cch = 0;
        gli.flags = 0;
        gli.yLine = lineEnd;
        UINT_PTR cch = (UINT_PTR)SendMessage(
            hwndView, EE_GET_LINEW, (WPARAM)&gli, (LPARAM)nullptr);
        ptEnd.y = lineEnd;
        ptEnd.x = (LONG_PTR)cch;
    }
    Editor_SetCaretPosEx(hwndView, POS_LOGICAL_W, &ptEnd, TRUE);

    // Build replacement text — strip null characters (c_str() would truncate)
    std::wstring replacement;
    replacement.reserve(newText.size());
    for (wchar_t ch : newText) {
        if (ch != L'\0') replacement += ch;
    }
    if (lineEnd + 1 < (int)totalLines && !replacement.empty() &&
        replacement.back() != L'\n') {
        replacement += L"\n";
    }

    // Replace selection with new text
    Editor_InsertW(hwndView, replacement.c_str(), false);

    // Invalidate cache so preview re-renders
    m_nLastHash = 0;
    m_sLastContent.clear();
}

// ============================================================================
// OnPreviewMermaidNodeEdited (M2)
//
// blockId is "mermaid-placeholder-N" (validated by the dispatcher).
// We refresh the live document, locate the Nth mermaid block, then ask
// MarkdownParser::ExtractFlowchartNodes to find the [labelStart, labelEnd)
// span for this nodeId. Replace it in-place, optionally promote to the
// quoted form when newLabel contains characters that would otherwise
// terminate the bracket pair, and feed the rebuilt block through the
// existing line-range writeback pipeline (OnPreviewTextEdited).
// ============================================================================
void CMermaidFrame::OnPreviewMermaidNodeEdited(HWND hwndView,
                                               const std::wstring& blockId,
                                               const std::wstring& nodeId,
                                               const std::wstring& newLabel)
{
    if (!hwndView || !IsWindow(hwndView)) return;
    if (blockId.size() <= 20) return; // "mermaid-placeholder-" + at least one digit

    // Block index = trailing integer of blockId.
    int blockIdx = 0;
    for (size_t i = 20; i < blockId.size(); ++i) {
        wchar_t c = blockId[i];
        if (c < L'0' || c > L'9') return;
        blockIdx = blockIdx * 10 + (c - L'0');
        if (blockIdx > 100000) return; // sanity cap
    }

    // Pull the freshest content from EmEditor — m_sLastContent may be stale
    // if the user typed in the editor between render and edit-commit.
    std::wstring content = MarkdownParser::GetDocumentContent(hwndView);
    auto blocks = MarkdownParser::ExtractMermaidBlocks(content);
    if (blockIdx < 0 || blockIdx >= (int)blocks.size()) return;
    const MermaidBlock& blk = blocks[blockIdx];

    auto nodes = MarkdownParser::ExtractFlowchartNodes(blk.code);
    const MermaidNodeRef* ref = nullptr;
    for (const auto& n : nodes) { if (n.nodeId == nodeId) { ref = &n; break; } }
    if (!ref) return;

    // Split block.code into lines (preserve content; we'll rejoin with \n).
    std::vector<std::wstring> lines;
    {
        size_t pos = 0;
        while (pos <= blk.code.size()) {
            size_t eol = blk.code.find(L'\n', pos);
            if (eol == std::wstring::npos) eol = blk.code.size();
            std::wstring l = blk.code.substr(pos, eol - pos);
            if (!l.empty() && l.back() == L'\r') l.pop_back();
            lines.push_back(std::move(l));
            if (eol >= blk.code.size()) break;
            pos = eol + 1;
        }
    }
    if (ref->lineOffsetInBlock < 0 || ref->lineOffsetInBlock >= (int)lines.size())
        return;
    std::wstring& targetLine = lines[ref->lineOffsetInBlock];
    if (ref->labelStart > ref->labelEnd || ref->labelEnd > targetLine.size())
        return;

    // Convert UI '\n' back to mermaid '<br/>'. Strip stray \r.
    std::wstring serialized;
    serialized.reserve(newLabel.size() + 8);
    for (wchar_t c : newLabel) {
        if (c == L'\r') continue;
        if (c == L'\n') serialized += L"<br/>";
        else            serialized += c;
    }

    // Decide whether the resulting label requires quoting. Bare labels
    // can't contain the matching close bracket, double-quote or pipe.
    wchar_t close = (ref->openBracket == L'[') ? L']'
                  : (ref->openBracket == L'{') ? L'}' : L')';
    bool needsQuote = false;
    for (wchar_t c : serialized) {
        if (c == close || c == L'"' || c == L'|') { needsQuote = true; break; }
    }

    std::wstring replacement;
    if (ref->isQuoted) {
        // Already inside quotes — escape any inner '"' to keep it parsable.
        replacement.reserve(serialized.size() + 4);
        for (wchar_t c : serialized) {
            if (c == L'"') replacement += L"#quot;"; // mermaid 11 entity
            else           replacement += c;
        }
    } else if (needsQuote) {
        replacement.reserve(serialized.size() + 2);
        replacement += L'"';
        for (wchar_t c : serialized) {
            if (c == L'"') replacement += L"#quot;";
            else           replacement += c;
        }
        replacement += L'"';
    } else {
        replacement = serialized;
    }

    // Splice into the source line.
    targetLine = targetLine.substr(0, ref->labelStart)
               + replacement
               + targetLine.substr(ref->labelEnd);

    // Rebuild the full mermaid block text — fences + body — to feed back
    // through the existing line-range edit path. ExtractMermaidBlocks
    // sets startLine = the ```mermaid line and endLine = the closing ```.
    std::wstring newBlock;
    // Reconstruct opening fence line by reading the actual document line —
    // keep any leading whitespace / language tag untouched.
    {
        UINT_PTR totalLines = (UINT_PTR)SendMessage(
            hwndView, EE_GET_LINES, (WPARAM)0, 0);
        if (blk.startLine < 0 || blk.startLine >= (int)totalLines) return;

        // Read a single line from EmEditor (mirrors MarkdownParser::GetDocumentContent).
        auto getLine = [&](int y) -> std::wstring {
            GET_LINE_INFO gli = {};
            gli.cch = 0;
            gli.flags = 0;
            gli.yLine = (UINT_PTR)y;
            UINT_PTR cch = (UINT_PTR)SendMessage(
                hwndView, EE_GET_LINEW, (WPARAM)&gli, (LPARAM)nullptr);
            std::wstring buf;
            if (cch == 0) return buf;
            buf.assign(cch, L'\0');
            gli.cch = cch;
            SendMessage(hwndView, EE_GET_LINEW, (WPARAM)&gli, (LPARAM)buf.data());
            while (!buf.empty() && buf.back() == L'\0') buf.pop_back();
            return buf;
        };
        std::wstring openFence = getLine(blk.startLine);
        std::wstring closeFence = (blk.endLine >= 0 && blk.endLine < (int)totalLines)
                                    ? getLine(blk.endLine)
                                    : L"```";
        newBlock += openFence; newBlock += L"\n";
        for (size_t i = 0; i < lines.size(); ++i) {
            newBlock += lines[i];
            newBlock += L"\n";
        }
        newBlock += closeFence;
    }

    // Reuse the existing line-range writeback so EmEditor undo, scroll
    // sync flags and m_*Last* invalidation all behave identically to the
    // plain-text edit path.
    OnPreviewTextEdited(hwndView, blk.startLine, blk.endLine, newBlock);
}

// ============================================================================
// OnPreviewMermaidBlockEdited (Phase 1 auto-correction)
//
// Whole-block replacement: JS computed and validated newSource, we just
// preserve the original ```mermaid / ``` fence lines and feed the full
// rebuilt block through OnPreviewTextEdited.
// ============================================================================
void CMermaidFrame::OnPreviewMermaidBlockEdited(HWND hwndView,
                                                const std::wstring& blockId,
                                                const std::wstring& newSource)
{
    if (!hwndView || !IsWindow(hwndView)) return;
    if (blockId.size() <= 20) return;

    int blockIdx = 0;
    for (size_t i = 20; i < blockId.size(); ++i) {
        wchar_t c = blockId[i];
        if (c < L'0' || c > L'9') return;
        blockIdx = blockIdx * 10 + (c - L'0');
        if (blockIdx > 100000) return;
    }

    // SEC-006: refuse any body that contains a fence-closing sequence —
    // a triple backtick inside the new mermaid source would prematurely
    // terminate the markdown code fence on writeback and corrupt the doc
    // structure. mermaid.parse on the JS side normally rejects such
    // input, but this is a defence-in-depth check we can do for free.
    if (newSource.find(L"```") != std::wstring::npos) return;

    std::wstring content = MarkdownParser::GetDocumentContent(hwndView);
    auto blocks = MarkdownParser::ExtractMermaidBlocks(content);
    if (blockIdx < 0 || blockIdx >= (int)blocks.size()) return;
    const MermaidBlock& blk = blocks[blockIdx];

    // Read the actual fence lines so we keep any leading whitespace / lang
    // tag the user wrote (e.g. `\t```mermaid` inside a list).
    auto getLine = [&](int y) -> std::wstring {
        GET_LINE_INFO gli = {};
        gli.cch = 0;
        gli.flags = 0;
        gli.yLine = (UINT_PTR)y;
        UINT_PTR cch = (UINT_PTR)SendMessage(
            hwndView, EE_GET_LINEW, (WPARAM)&gli, (LPARAM)nullptr);
        std::wstring buf;
        if (cch == 0) return buf;
        buf.assign(cch, L'\0');
        gli.cch = cch;
        SendMessage(hwndView, EE_GET_LINEW, (WPARAM)&gli, (LPARAM)buf.data());
        while (!buf.empty() && buf.back() == L'\0') buf.pop_back();
        return buf;
    };
    UINT_PTR totalLines = (UINT_PTR)SendMessage(
        hwndView, EE_GET_LINES, (WPARAM)0, 0);
    if (blk.startLine < 0 || blk.startLine >= (int)totalLines) return;
    std::wstring openFence  = getLine(blk.startLine);
    std::wstring closeFence = (blk.endLine >= 0 && blk.endLine < (int)totalLines)
                                ? getLine(blk.endLine)
                                : L"```";

    // Normalise the body — strip an optional trailing newline so we don't
    // double up on the join below; we add exactly one between each line.
    std::wstring body = newSource;
    if (!body.empty() && body.back() == L'\n') body.pop_back();
    if (!body.empty() && body.back() == L'\r') body.pop_back();

    std::wstring newBlock;
    newBlock.reserve(openFence.size() + body.size() + closeFence.size() + 4);
    newBlock += openFence;
    newBlock += L'\n';
    newBlock += body;
    newBlock += L'\n';
    newBlock += closeFence;

    OnPreviewTextEdited(hwndView, blk.startLine, blk.endLine, newBlock);
}

// ============================================================================
// IsDarkMode
// ============================================================================
bool CMermaidFrame::IsDarkMode(HWND hwndView) const
{
    if (!hwndView || !IsWindow(hwndView))
        return false;

    LRESULT result = Editor_Info(hwndView, EI_IS_VERY_DARK, 0);
    if (result == TRUE)
        return true;

    if (result == NOT_SUPPORTED)
        return false;

    LRESULT bgColor = Editor_Info(hwndView, EI_GET_BAR_BACK_COLOR, 0);
    if (bgColor != 0) {
        COLORREF cr = static_cast<COLORREF>(bgColor);
        int r = GetRValue(cr);
        int g = GetGValue(cr);
        int b = GetBValue(cr);
        int luminance = (r * 299 + g * 587 + b * 114) / 1000;
        return luminance < 128;
    }

    return false;
}

// ============================================================================
// Registry persistence
//
// `iSchemaVersion` is the anchor for future migrations (audit v2 LOW-4.5).
// Bump when key names or value semantics change incompatibly; LoadSettings
// can then branch on the read value to migrate from older layouts.
// ============================================================================
constexpr int kSettingsSchemaVersion = 1;

void CMermaidFrame::LoadSettings()
{
    int schemaVersion = GetProfileInt(L"iSchemaVersion", 0);
    (void)schemaVersion; // currently no migration paths; placeholder for v2+

    m_iBarPos = GetProfileInt(L"iPos", 2);
    if (m_iBarPos < 0 || m_iBarPos > 3)
        m_iBarPos = 2; // Clamp to valid Custom Bar positions
    m_bDarkMode = GetProfileInt(L"iDarkMode", 0) != 0;
    m_bDarkModeOverride = GetProfileInt(L"iDarkModeOverride", 0) != 0;
    m_iFontSize = GetProfileInt(L"iFontSize", 14);
    if (m_iFontSize < 8 || m_iFontSize > 32)
        m_iFontSize = 14;
}

void CMermaidFrame::SaveSettings()
{
    WriteProfileInt(L"iSchemaVersion", kSettingsSchemaVersion);
    WriteProfileInt(L"iPos", m_iBarPos);
    WriteProfileInt(L"iDarkMode", m_bDarkMode ? 1 : 0);
    WriteProfileInt(L"iDarkModeOverride", m_bDarkModeOverride ? 1 : 0);
    WriteProfileInt(L"iFontSize", m_iFontSize);
}
