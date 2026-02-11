#pragma once

#include <windows.h>
#include <string>
#include <vector>
#include <functional>

struct MermaidRenderResult {
    std::wstring id;
    std::wstring svg;    // Empty on error
    std::wstring error;  // Empty on success
};

class BunRenderer {
public:
    BunRenderer();
    ~BunRenderer();

    // Start the persistent Bun process. Returns true on success.
    bool Start();

    // Stop the Bun process.
    void Stop();

    // Is the Bun process running and ready?
    bool IsReady() const { return m_bReady; }

    // Render mermaid blocks to SVG. Blocks the calling thread briefly.
    // theme: "default" or "dark"
    std::vector<MermaidRenderResult> RenderBlocks(
        const std::vector<std::pair<std::wstring, std::wstring>>& blocks, // {id, code}
        const std::wstring& theme);

private:
    // Send a line of JSON to Bun's stdin
    bool SendLine(const std::string& json);

    // Read a line of JSON from Bun's stdout (with timeout)
    std::string ReadLine(DWORD timeoutMs = 5000);

    // Find Bun executable path
    std::wstring FindBunPath() const;

    // Get renderer script path
    std::wstring GetRendererPath() const;

    // Ensure bun-renderer directory is set up
    bool EnsureSetup();

    // Convert wstring to UTF-8
    static std::string WtoU8(const std::wstring& ws);

    // Convert UTF-8 to wstring
    static std::wstring U8toW(const std::string& s);

    // Escape a string for JSON value
    static std::string JsonEscape(const std::string& s);

    HANDLE m_hProcess = nullptr;
    HANDLE m_hStdinWrite = nullptr;
    HANDLE m_hStdoutRead = nullptr;
    bool m_bReady = false;
    std::string m_readBuffer;
};
