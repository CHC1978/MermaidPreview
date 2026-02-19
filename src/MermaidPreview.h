#pragma once

// NOTE: etlframe.h (with ETL_FRAME_CLASS_NAME defined) must be #included
// before this file, so that CETLFrame<T> template is fully defined.

#include <string>
#include <memory>
#include <vector>
#include <future>
#include "resource.h"

class WebView2Manager;
class BunRenderer;

class CMermaidFrame : public CETLFrame<CMermaidFrame> {
public:
    // --- Compile-time configuration required by CETLFrame ---
    enum {
        _IDS_NAME                   = IDS_PLUGIN_NAME,
        _IDS_MENU                   = IDS_MENU_TEXT,
        _IDS_STATUS                 = IDS_STATUS_TEXT,
        _IDS_VER                    = IDS_VERSION,

        _IDB_BITMAP                 = IDB_BITMAP,
        _IDB_TRUE_16_DEFAULT        = IDB_TRUE_16_DEFAULT,
        _IDB_TRUE_16_BW             = IDB_TRUE_16_BW,
        _IDB_TRUE_16_HOT            = IDB_TRUE_16_HOT,
        _IDB_16C_24                 = IDB_16C_24,
        _IDB_TRUE_24_DEFAULT        = IDB_TRUE_24_DEFAULT,
        _IDB_TRUE_24_BW             = IDB_TRUE_24_BW,
        _IDB_TRUE_24_HOT            = IDB_TRUE_24_HOT,
        _MASK_TRUE_COLOR            = 0,

        _USE_LOC_DLL                = LOC_USE_PLUGIN_DLL,
        _SUPPORT_EE_PRO             = TRUE,
        _USE_CUSTOM_BAR             = TRUE,
        _ALLOW_OPEN_SAME_GROUP      = FALSE,
        _ALLOW_MULTIPLE_INSTANCES   = FALSE,
    };

    // --- EmEditor callbacks ---
    void OnCommand(HWND hwndView);
    void OnEvents(HWND hwndView, UINT nEvent, LPARAM lParam);
    BOOL QueryStatus(HWND hwndView, LPBOOL pbChecked);
    BOOL QueryUninstall(HWND /*hwnd*/) { return TRUE; }
    BOOL SetUninstall(HWND /*hwnd*/, LPWSTR /*pszCmd*/, LPWSTR /*pszParam*/) { return FALSE; }
    BOOL QueryProperties(HWND /*hwnd*/) { return FALSE; }
    BOOL SetProperties(HWND /*hwnd*/) { return FALSE; }
    BOOL PreTranslateMessage(HWND /*hwnd*/, MSG* /*pMsg*/) { return FALSE; }
    BOOL UseDroppedFiles(HWND /*hwnd*/) { return FALSE; }
    BOOL DisableAutoComplete(HWND /*hwnd*/) { return FALSE; }
    LRESULT UserMessage(HWND /*hwnd*/, WPARAM /*wParam*/, LPARAM /*lParam*/) { return 0; }

private:
    // --- Custom Bar management ---
    void OpenCustomBar(HWND hwndView, std::wstring prefetchedContent = {});
    void CloseCustomBar(HWND hwndView);
    void OnCustomBarClosed(HWND hwndView, LPARAM lParam);

    // --- Preview logic ---
    void UpdatePreview(HWND hwndView);
    bool IsDarkMode(HWND hwndView) const;

    // --- Bun renderer ---
    void EnsureBunRenderer();

    // --- Edit callback from WebView2 ---
    void OnPreviewTextEdited(HWND hwndView, int lineStart, int lineEnd,
                             const std::wstring& newText);

    // --- Scroll sync ---
    void SyncScrollToPreview(HWND hwndView);
    void OnPreviewScrolled(HWND hwndView, int line);
    void OnPreviewNavigate(HWND hwndView, int line);

    // --- Open relative-path file link ---
    void OnOpenFileLink(HWND hwndView, const std::wstring& relativePath);

    // --- Auto-open detection ---
    bool IsMarkdownFile(HWND hwndView) const;
    bool HasMermaidBlocks(HWND hwndView) const;
    void TryAutoOpen(HWND hwndView);
    void TryAutoClose(HWND hwndView);

    // --- Registry ---
    void LoadSettings();
    void SaveSettings();

    // --- Window procedure for the custom bar host ---
    static LRESULT CALLBACK HostWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);

    // --- State ---
    HWND                            m_hwndHost = nullptr;
    HWND                            m_hwndParking = nullptr;  // Hidden parking window for WebView2
    HWND                            m_hWndLastView = nullptr;
    UINT                            m_nBarID = 0;
    bool                            m_bVisible = false;
    bool                            m_bParked = false;        // WebView2 is parked (hidden, not destroyed)
    bool                            m_bAutoOpened = false;
    std::unique_ptr<WebView2Manager> m_pWebView;
    std::shared_ptr<BunRenderer>    m_pBunRenderer;
    std::wstring                    m_sLastContent;
    size_t                          m_nLastHash = 0;
    bool                            m_bDarkMode = false;
    bool                            m_bDarkModeOverride = false; // User manual override
    bool                            m_bSyncFromEditor = false;   // Anti-feedback: Editor→Preview
    bool                            m_bSyncFromPreview = false;  // Anti-feedback: Preview→Editor
    bool                            m_bBunAvailable = false;
    std::future<bool>               m_bunStartFuture;

    // Optimization 3: Pre-fetched HTML (prepared while WebView2 initializes)
    std::wstring                    m_sPrefetchedHtml;
    bool                            m_bHasPrefetch = false;

    // --- Settings ---
    int                             m_iBarPos = 2;
    int                             m_iFontSize = 14;
};
