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

// ============================================================================
// Custom Bar host window class name
// ============================================================================
static const wchar_t* const kHostClassName = L"MermaidPreviewHost";
static bool s_bHostClassRegistered = false;

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
        break;
    }
    case WM_DESTROY:
        return 0;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
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
        return;
    }

    if (nEvent & EVENT_CLOSE_FRAME) {
        SaveSettings();
        if (m_bVisible) {
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
            m_bVisible = false;
            m_bAutoOpened = false;
        }
        // Wait for async Bun startup to complete, then stop
        if (m_bunStartFuture.valid()) {
            m_bunStartFuture.wait();
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

    if (nEvent & EVENT_DOC_SEL_CHANGED) {
        m_hWndLastView = hwndView;
        if (!m_bVisible) {
            TryAutoOpen(hwndView);
        } else {
            m_nLastHash = 0;
            m_sLastContent.clear();
            if (m_bAutoOpened) {
                TryAutoClose(hwndView);
            }
            if (m_bVisible) {
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
        m_pBunRenderer = std::make_unique<BunRenderer>();
    }

    // Launch Bun startup in background thread to avoid freezing UI
    BunRenderer* pRenderer = m_pBunRenderer.get();
    m_bunStartFuture = std::async(std::launch::async, [pRenderer]() {
        return pRenderer->Start();
    });
}

// ============================================================================
// OpenCustomBar
// ============================================================================
void CMermaidFrame::OpenCustomBar(HWND hwndView)
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
    if (!m_bDarkModeOverride) {
        m_bDarkMode = IsDarkMode(hwndView);
    }

    // Try to start Bun renderer in background
    EnsureBunRenderer();

    // Initialize WebView2
    m_pWebView = std::make_unique<WebView2Manager>();
    m_pWebView->Initialize(m_hwndHost, [this, hwndView]() {
        if (m_pWebView) {
            m_pWebView->SetTheme(m_bDarkMode);

            // Register WebMessage handler for editing
            m_pWebView->SetEditCallback([this, hwndView](int lineStart, int lineEnd, const std::wstring& newText) {
                OnPreviewTextEdited(hwndView, lineStart, lineEnd, newText);
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

            // Optimization 3: Use pre-fetched content if available
            if (m_bHasPrefetch) {
                m_pWebView->RenderContent(m_sPrefetchedHtml, m_bDarkMode);
                SyncScrollToPreview(hwndView);
                m_bHasPrefetch = false;
                m_sPrefetchedHtml.clear();
            } else {
                UpdatePreview(hwndView);
            }
        }
    });

    // Optimization 3: Pre-fetch document content while WebView2 initializes async
    // This overlaps content preparation with the ~800-1500ms WebView2 startup
    {
        std::wstring content = MarkdownParser::GetDocumentContent(hwndView);
        std::hash<std::wstring> hasher;
        m_nLastHash = hasher(content);
        m_sLastContent = content;
        m_sPrefetchedHtml = MarkdownParser::ConvertToHtml(content);
        m_bHasPrefetch = true;
    }

    m_hWndLastView = hwndView;
}

// ============================================================================
// CloseCustomBar
// ============================================================================
void CMermaidFrame::CloseCustomBar(HWND /*hwndView*/)
{
    if (!m_bVisible)
        return;

    if (m_hwndHost) {
        KillTimer(m_hwndHost, IDT_DEBOUNCE);
        KillTimer(m_hwndHost, IDT_SCROLL_SYNC);
        KillTimer(m_hwndHost, IDT_SYNC_RESET_E2P);
        KillTimer(m_hwndHost, IDT_SYNC_RESET_P2E);
    }

    if (m_pWebView) {
        m_pWebView->Destroy();
        m_pWebView.reset();
    }

    if (m_nBarID != 0) {
        Editor_CustomBarClose(m_hWnd, m_nBarID);
        m_nBarID = 0;
    }

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
    }

    if (m_pWebView) {
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

void CMermaidFrame::TryAutoOpen(HWND hwndView)
{
    if (m_bVisible) return;
    if (IsMarkdownFile(hwndView) && HasMermaidBlocks(hwndView)) {
        OpenCustomBar(hwndView);
        m_bAutoOpened = true;
    }
}

void CMermaidFrame::TryAutoClose(HWND hwndView)
{
    if (!m_bVisible || !m_bAutoOpened) return;
    if (!IsMarkdownFile(hwndView) || !HasMermaidBlocks(hwndView)) {
        CloseCustomBar(hwndView);
        m_bAutoOpened = false;
    }
}

// ============================================================================
// UpdatePreview - Hybrid: C++ markdown + Bun mermaid SVG (fallback: WebView2 JS)
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

    // Try Bun renderer for mermaid blocks (pre-render to SVG)
    if (m_bBunAvailable && m_pBunRenderer && m_pBunRenderer->IsReady()) {
        auto mermaidBlocks = MarkdownParser::ExtractMermaidBlocks(content);

        if (!mermaidBlocks.empty()) {
            std::vector<std::pair<std::wstring, std::wstring>> bunBlocks;
            for (size_t i = 0; i < mermaidBlocks.size(); i++) {
                bunBlocks.push_back({
                    L"mermaid-placeholder-" + std::to_wstring(i),
                    mermaidBlocks[i].code
                });
            }

            std::wstring theme = m_bDarkMode ? L"dark" : L"default";
            auto results = m_pBunRenderer->RenderBlocks(bunBlocks, theme);

            // Replace mermaid placeholders with pre-rendered SVG
            for (auto& r : results) {
                std::wstring placeholder = L"data-mermaid-id=\"" + r.id + L"\"";
                size_t pos = html.find(placeholder);
                if (pos == std::wstring::npos) continue;

                // Find the container div boundaries
                size_t divStart = html.rfind(L"<div ", pos);
                size_t divEnd = html.find(L"</div>", pos);
                if (divStart == std::wstring::npos || divEnd == std::wstring::npos) continue;
                divEnd += 6; // Include </div>

                // Extract data-mermaid-src and line tracking from original div
                std::wstring origDiv = html.substr(divStart, divEnd - divStart);
                std::wstring dataSrc;
                size_t srcAttr = origDiv.find(L"data-mermaid-src=\"");
                if (srcAttr != std::wstring::npos) {
                    srcAttr += 18;
                    size_t srcEnd = origDiv.find(L'"', srcAttr);
                    if (srcEnd != std::wstring::npos)
                        dataSrc = origDiv.substr(srcAttr, srcEnd - srcAttr);
                }

                // Preserve data-line-start and data-line-end for editing
                std::wstring dataLineStart, dataLineEnd;
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
                if (!dataSrc.empty())
                    attrs += L" data-mermaid-src=\"" + dataSrc + L"\"";
                if (!dataLineStart.empty())
                    attrs += L" data-line-start=\"" + dataLineStart + L"\"";
                if (!dataLineEnd.empty())
                    attrs += L" data-line-end=\"" + dataLineEnd + L"\"";

                if (!r.svg.empty()) {
                    std::wstring svgDiv = L"<div " + attrs + L">" + r.svg + L"</div>";
                    html = html.substr(0, divStart) + svgDiv + html.substr(divEnd);
                } else if (!r.error.empty()) {
                    std::wstring errDiv = L"<div " + attrs + L"><div class=\"mermaid-error\">Mermaid error: "
                        + r.error + L"</div></div>";
                    html = html.substr(0, divStart) + errDiv + html.substr(divEnd);
                }
            }
        }
    }
    // If Bun not available, mermaid placeholders remain in HTML
    // and WebView2's mermaid.js will handle them (fallback)

    m_pWebView->RenderContent(html, m_bDarkMode);

    // After content update, sync preview scroll to editor's current position
    SyncScrollToPreview(hwndView);
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

    // Clamp lineEnd
    if (lineEnd >= (int)totalLines)
        lineEnd = (int)totalLines - 1;

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

    // Build replacement text (add newline at end if replacing full lines)
    std::wstring replacement = newText;
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
// ============================================================================
void CMermaidFrame::LoadSettings()
{
    m_iBarPos = GetProfileInt(L"iPos", 2);
    m_bDarkMode = GetProfileInt(L"iDarkMode", 0) != 0;
    m_bDarkModeOverride = GetProfileInt(L"iDarkModeOverride", 0) != 0;
}

void CMermaidFrame::SaveSettings()
{
    WriteProfileInt(L"iPos", m_iBarPos);
    WriteProfileInt(L"iDarkMode", m_bDarkMode ? 1 : 0);
    WriteProfileInt(L"iDarkModeOverride", m_bDarkModeOverride ? 1 : 0);
}
