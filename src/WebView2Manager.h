#pragma once

#include <windows.h>
#include <wrl.h>
#include <WebView2.h>
#include <string>
#include <vector>
#include <functional>

class WebView2Manager {
public:
    WebView2Manager();
    ~WebView2Manager();

    // Initialize WebView2 in the given parent window.
    // onReady is called when the WebView2 is fully initialized.
    HRESULT Initialize(HWND hwndParent, std::function<void()> onReady);

    // Render pre-parsed HTML content (C++ side already converted markdown to HTML).
    // Mermaid placeholders are rendered incrementally by comparing with previous state.
    void RenderContent(const std::wstring& htmlContent, bool darkMode);

    // Switch light/dark theme
    void SetTheme(bool darkMode);

    // Resize the WebView2 to match parent bounds
    void Resize(RECT bounds);

    // Clear all content
    void Clear();

    // Destroy and release all COM resources
    void Destroy();

    // Is the WebView2 ready to render?
    bool IsReady() const { return m_bReady; }

    // Parking window support: move WebView2 to a hidden window without destroying
    void Park(HWND hwndParking);

    // Reparent WebView2 to a new host window (fast reopen)
    HRESULT Reparent(HWND hwndNewParent);

    // Check if the controller has been created (for park/destroy decisions)
    bool HasController() const { return m_controller != nullptr; }

    // Set callback for text editing from preview panel
    using EditCallback = std::function<void(int lineStart, int lineEnd, const std::wstring& newText)>;
    void SetEditCallback(EditCallback callback);

    // Set callback for theme change from context menu
    using ThemeCallback = std::function<void(bool dark)>;
    void SetThemeCallback(ThemeCallback callback);

    // Scroll sync: scroll preview to match editor line
    void ScrollToLine(int line);

    // Set callback for scroll sync from preview
    using ScrollCallback = std::function<void(int line)>;
    void SetScrollCallback(ScrollCallback callback);

    // Set callback for click-to-navigate from preview (scroll + caret)
    using NavigateCallback = std::function<void(int line)>;
    void SetNavigateCallback(NavigateCallback callback);

private:
    // Build the HTML page shell (CSS + mermaid rendering JS)
    // mermaid.js is loaded from local extracted file
    std::wstring BuildHtmlPage() const;

    // Extract mermaid.min.js from DLL resource to local directory
    bool ExtractMermaidJs();

    // Get the local resource directory path
    std::wstring GetResourceDir() const;

    // Escape a string for safe embedding in JavaScript
    static std::wstring EscapeForJS(const std::wstring& input);

    Microsoft::WRL::ComPtr<ICoreWebView2Environment> m_env;
    Microsoft::WRL::ComPtr<ICoreWebView2Controller> m_controller;
    Microsoft::WRL::ComPtr<ICoreWebView2>           m_webview;
    HWND m_hwndParent = nullptr;
    bool m_bReady = false;
    bool m_bMessageHandlerRegistered = false;
    std::function<void()> m_onReady;

    // COM event tokens for proper unsubscription in Destroy()
    EventRegistrationToken m_navCompletedToken = {};
    EventRegistrationToken m_webMessageToken = {};

    // Cancellation: set by Destroy() to prevent async callbacks from running
    bool m_bDestroyed = false;

    // Local resource directory (mermaid.min.js location)
    std::wstring m_resourceDir;

    // Pending render request (if called before ready)
    std::wstring m_pendingHtml;
    bool m_pendingDarkMode = false;
    bool m_hasPendingRender = false;

    // Callbacks from preview panel
    EditCallback m_editCallback;
    ThemeCallback m_themeCallback;
    ScrollCallback m_scrollCallback;
    NavigateCallback m_navigateCallback;
};
