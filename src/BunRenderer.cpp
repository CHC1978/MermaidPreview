#include "BunRenderer.h"
#include <shlobj.h>

// Get the DLL's HMODULE
extern HINSTANCE EEGetInstanceHandle();

BunRenderer::BunRenderer() = default;

BunRenderer::~BunRenderer()
{
    Stop();
}

// ============================================================================
// WtoU8 / U8toW - encoding conversions
// ============================================================================
std::string BunRenderer::WtoU8(const std::wstring& ws)
{
    if (ws.empty()) return {};
    int len = WideCharToMultiByte(CP_UTF8, 0, ws.c_str(), (int)ws.size(), nullptr, 0, nullptr, nullptr);
    std::string s(len, '\0');
    WideCharToMultiByte(CP_UTF8, 0, ws.c_str(), (int)ws.size(), s.data(), len, nullptr, nullptr);
    return s;
}

std::wstring BunRenderer::U8toW(const std::string& s)
{
    if (s.empty()) return {};
    int len = MultiByteToWideChar(CP_UTF8, 0, s.c_str(), (int)s.size(), nullptr, 0);
    std::wstring ws(len, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, s.c_str(), (int)s.size(), ws.data(), len);
    return ws;
}

// ============================================================================
// JsonEscape
// ============================================================================
std::string BunRenderer::JsonEscape(const std::string& s)
{
    std::string out;
    out.reserve(s.size() + s.size() / 4);
    for (char ch : s) {
        switch (ch) {
        case '"':  out += "\\\""; break;
        case '\\': out += "\\\\"; break;
        case '\n': out += "\\n";  break;
        case '\r': out += "\\r";  break;
        case '\t': out += "\\t";  break;
        default:
            if (static_cast<unsigned char>(ch) < 0x20) {
                char buf[8];
                sprintf_s(buf, "\\u%04x", (unsigned char)ch);
                out += buf;
            } else {
                out += ch;
            }
        }
    }
    return out;
}

// ============================================================================
// FindBunPath - locate bun.exe
// ============================================================================
std::wstring BunRenderer::FindBunPath() const
{
    // Check common locations
    WCHAR userProfile[MAX_PATH] = {};
    if (SUCCEEDED(SHGetFolderPathW(nullptr, CSIDL_PROFILE, nullptr, 0, userProfile))) {
        std::wstring bunPath = userProfile;
        bunPath += L"\\.bun\\bin\\bun.exe";
        if (GetFileAttributesW(bunPath.c_str()) != INVALID_FILE_ATTRIBUTES)
            return bunPath;
    }

    // Try PATH
    WCHAR pathBuf[MAX_PATH] = {};
    if (SearchPathW(nullptr, L"bun.exe", nullptr, MAX_PATH, pathBuf, nullptr))
        return pathBuf;

    return L"";
}

// ============================================================================
// GetRendererPath - path to the renderer.ts script
// ============================================================================
std::wstring BunRenderer::GetRendererPath() const
{
    // The renderer script is alongside the DLL in bun-renderer/
    WCHAR dllPath[MAX_PATH] = {};
    GetModuleFileNameW(EEGetInstanceHandle(), dllPath, MAX_PATH);

    // Remove filename to get directory
    std::wstring dir = dllPath;
    size_t lastSlash = dir.find_last_of(L"\\/");
    if (lastSlash != std::wstring::npos)
        dir = dir.substr(0, lastSlash);

    return dir + L"\\bun-renderer\\renderer.ts";
}

// ============================================================================
// EnsureSetup - verify bun-renderer directory and node_modules exist
// ============================================================================
bool BunRenderer::EnsureSetup()
{
    std::wstring rendererPath = GetRendererPath();
    if (GetFileAttributesW(rendererPath.c_str()) == INVALID_FILE_ATTRIBUTES)
        return false;

    // Check if node_modules exists
    std::wstring dir = rendererPath.substr(0, rendererPath.find_last_of(L"\\/"));
    std::wstring nodeModules = dir + L"\\node_modules";
    if (GetFileAttributesW(nodeModules.c_str()) == INVALID_FILE_ATTRIBUTES) {
        // Run bun install
        std::wstring bunPath = FindBunPath();
        if (bunPath.empty()) return false;

        std::wstring cmd = L"\"" + bunPath + L"\" install";

        STARTUPINFOW si = {};
        si.cb = sizeof(si);
        si.dwFlags = STARTF_USESHOWWINDOW;
        si.wShowWindow = SW_HIDE;

        PROCESS_INFORMATION pi = {};
        BOOL ok = CreateProcessW(
            nullptr, const_cast<LPWSTR>(cmd.c_str()),
            nullptr, nullptr, FALSE, CREATE_NO_WINDOW,
            nullptr, dir.c_str(), &si, &pi);

        if (!ok) return false;

        WaitForSingleObject(pi.hProcess, 60000); // 60s timeout
        CloseHandle(pi.hProcess);
        CloseHandle(pi.hThread);

        // Verify
        if (GetFileAttributesW(nodeModules.c_str()) == INVALID_FILE_ATTRIBUTES)
            return false;
    }

    return true;
}

// ============================================================================
// Start - spawn the persistent Bun process
// ============================================================================
bool BunRenderer::Start()
{
    if (m_bReady)
        return true;

    // Find bun
    std::wstring bunPath = FindBunPath();
    if (bunPath.empty())
        return false;

    // Ensure setup
    if (!EnsureSetup())
        return false;

    std::wstring rendererPath = GetRendererPath();
    std::wstring rendererDir = rendererPath.substr(0, rendererPath.find_last_of(L"\\/"));

    // Create pipes for stdin/stdout
    SECURITY_ATTRIBUTES sa = {};
    sa.nLength = sizeof(sa);
    sa.bInheritHandle = TRUE;
    sa.lpSecurityDescriptor = nullptr;

    HANDLE hStdinRead = nullptr, hStdinWrite = nullptr;
    HANDLE hStdoutRead = nullptr, hStdoutWrite = nullptr;

    if (!CreatePipe(&hStdinRead, &hStdinWrite, &sa, 0))
        return false;
    if (!CreatePipe(&hStdoutRead, &hStdoutWrite, &sa, 0)) {
        CloseHandle(hStdinRead);
        CloseHandle(hStdinWrite);
        return false;
    }

    // Ensure our end of the pipes are not inherited
    SetHandleInformation(hStdinWrite, HANDLE_FLAG_INHERIT, 0);
    SetHandleInformation(hStdoutRead, HANDLE_FLAG_INHERIT, 0);

    // Build command line: bun run renderer.ts
    std::wstring cmd = L"\"" + bunPath + L"\" run \"" + rendererPath + L"\"";

    STARTUPINFOW si = {};
    si.cb = sizeof(si);
    si.dwFlags = STARTF_USESTDHANDLES | STARTF_USESHOWWINDOW;
    si.hStdInput = hStdinRead;
    si.hStdOutput = hStdoutWrite;
    si.hStdError = hStdoutWrite; // Merge stderr to stdout
    si.wShowWindow = SW_HIDE;

    PROCESS_INFORMATION pi = {};
    BOOL ok = CreateProcessW(
        nullptr, const_cast<LPWSTR>(cmd.c_str()),
        nullptr, nullptr, TRUE, CREATE_NO_WINDOW,
        nullptr, rendererDir.c_str(), &si, &pi);

    // Close child-side handles (parent doesn't need them)
    CloseHandle(hStdinRead);
    CloseHandle(hStdoutWrite);

    if (!ok) {
        CloseHandle(hStdinWrite);
        CloseHandle(hStdoutRead);
        return false;
    }

    m_hProcess = pi.hProcess;
    CloseHandle(pi.hThread);
    m_hStdinWrite = hStdinWrite;
    m_hStdoutRead = hStdoutRead;

    // Wait for "ready" message (up to 5 seconds)
    std::string line = ReadLine(5000);
    if (line.find("\"ready\"") != std::string::npos) {
        m_bReady = true;
        return true;
    }

    // Failed to start
    Stop();
    return false;
}

// ============================================================================
// Stop - terminate the Bun process
// ============================================================================
void BunRenderer::Stop()
{
    m_bReady = false;

    if (m_hStdinWrite) {
        CloseHandle(m_hStdinWrite);
        m_hStdinWrite = nullptr;
    }
    if (m_hStdoutRead) {
        CloseHandle(m_hStdoutRead);
        m_hStdoutRead = nullptr;
    }
    if (m_hProcess) {
        TerminateProcess(m_hProcess, 0);
        WaitForSingleObject(m_hProcess, 3000);
        CloseHandle(m_hProcess);
        m_hProcess = nullptr;
    }
    m_readBuffer.clear();
}

// ============================================================================
// SendLine - write JSON line to stdin pipe
// ============================================================================
bool BunRenderer::SendLine(const std::string& json)
{
    if (!m_hStdinWrite) return false;
    std::string line = json + "\n";
    DWORD written = 0;
    return WriteFile(m_hStdinWrite, line.data(), (DWORD)line.size(), &written, nullptr) && written == line.size();
}

// ============================================================================
// ReadLine - read a line from stdout pipe (with timeout)
// ============================================================================
std::string BunRenderer::ReadLine(DWORD timeoutMs)
{
    if (!m_hStdoutRead) return "";

    DWORD startTime = GetTickCount();
    char buf[4096];

    while (true) {
        // Check if we already have a line in the buffer
        size_t nl = m_readBuffer.find('\n');
        if (nl != std::string::npos) {
            std::string line = m_readBuffer.substr(0, nl);
            m_readBuffer = m_readBuffer.substr(nl + 1);
            // Trim \r
            if (!line.empty() && line.back() == '\r')
                line.pop_back();
            return line;
        }

        // Check timeout
        DWORD elapsed = GetTickCount() - startTime;
        if (elapsed >= timeoutMs)
            return "";

        // Check if data available
        DWORD available = 0;
        if (!PeekNamedPipe(m_hStdoutRead, nullptr, 0, nullptr, &available, nullptr))
            return "";

        if (available == 0) {
            // Check if process is still alive
            DWORD exitCode = 0;
            if (!GetExitCodeProcess(m_hProcess, &exitCode) || exitCode != STILL_ACTIVE)
                return "";
            Sleep(10);
            continue;
        }

        DWORD toRead = (available < sizeof(buf)) ? available : sizeof(buf);
        DWORD bytesRead = 0;
        if (!ReadFile(m_hStdoutRead, buf, toRead, &bytesRead, nullptr) || bytesRead == 0)
            return "";

        m_readBuffer.append(buf, bytesRead);
    }
}

// ============================================================================
// RenderBlocks - send mermaid code to Bun, get SVG back
// ============================================================================
std::vector<MermaidRenderResult> BunRenderer::RenderBlocks(
    const std::vector<std::pair<std::wstring, std::wstring>>& blocks,
    const std::wstring& theme)
{
    std::vector<MermaidRenderResult> results;

    if (!m_bReady || blocks.empty())
        return results;

    // Build JSON request
    std::string json = "{\"type\":\"render\",\"blocks\":[";
    for (size_t i = 0; i < blocks.size(); i++) {
        if (i > 0) json += ",";
        json += "{\"id\":\"" + JsonEscape(WtoU8(blocks[i].first)) + "\",";
        json += "\"code\":\"" + JsonEscape(WtoU8(blocks[i].second)) + "\"}";
    }
    json += "],\"theme\":\"" + JsonEscape(WtoU8(theme)) + "\"}";

    if (!SendLine(json))
        return results;

    // Read response (timeout based on block count)
    DWORD timeout = 5000 + (DWORD)blocks.size() * 3000;
    std::string response = ReadLine(timeout);

    if (response.empty())
        return results;

    // Simple JSON parsing for the response
    // Format: {"type":"result","results":[{"id":"...","svg":"...","error":null},...]}
    // We use a basic approach since we control both sides of the protocol

    // Find "results" array
    size_t resultsStart = response.find("\"results\":[");
    if (resultsStart == std::string::npos)
        return results;

    resultsStart += 11; // Skip "results":[

    // Parse each result object
    size_t pos = resultsStart;
    while (pos < response.size()) {
        size_t objStart = response.find('{', pos);
        if (objStart == std::string::npos) break;

        // Find id
        size_t idKey = response.find("\"id\":\"", objStart);
        if (idKey == std::string::npos) break;
        idKey += 6;
        size_t idEnd = response.find('"', idKey);
        std::string id = response.substr(idKey, idEnd - idKey);

        MermaidRenderResult r;
        r.id = U8toW(id);

        // Check for svg
        size_t svgKey = response.find("\"svg\":", idEnd);
        if (svgKey != std::string::npos) {
            size_t svgValStart = svgKey + 6;
            // Skip whitespace
            while (svgValStart < response.size() && response[svgValStart] == ' ') svgValStart++;

            if (response[svgValStart] == '"') {
                // SVG string value - find the matching close quote (handle escaped quotes)
                svgValStart++;
                std::string svgStr;
                size_t si = svgValStart;
                while (si < response.size()) {
                    if (response[si] == '\\' && si + 1 < response.size()) {
                        char next = response[si + 1];
                        if (next == '"') { svgStr += '"'; si += 2; }
                        else if (next == '\\') { svgStr += '\\'; si += 2; }
                        else if (next == 'n') { svgStr += '\n'; si += 2; }
                        else if (next == 'r') { svgStr += '\r'; si += 2; }
                        else if (next == 't') { svgStr += '\t'; si += 2; }
                        else if (next == '/') { svgStr += '/'; si += 2; }
                        else { svgStr += response[si]; si++; }
                    } else if (response[si] == '"') {
                        break;
                    } else {
                        svgStr += response[si];
                        si++;
                    }
                }
                r.svg = U8toW(svgStr);
                pos = si + 1;
            } else if (response.substr(svgValStart, 4) == "null") {
                pos = svgValStart + 4;
            }
        }

        // Check for error
        size_t errKey = response.find("\"error\":", pos);
        if (errKey != std::string::npos) {
            size_t errValStart = errKey + 8;
            while (errValStart < response.size() && response[errValStart] == ' ') errValStart++;
            if (response[errValStart] == '"') {
                errValStart++;
                size_t errEnd = response.find('"', errValStart);
                if (errEnd != std::string::npos) {
                    r.error = U8toW(response.substr(errValStart, errEnd - errValStart));
                    pos = errEnd + 1;
                }
            } else if (response.substr(errValStart, 4) == "null") {
                pos = errValStart + 4;
            }
        }

        results.push_back(std::move(r));

        // Find next object or end of array
        size_t nextComma = response.find(',', pos);
        size_t nextBracket = response.find(']', pos);
        if (nextBracket != std::string::npos &&
            (nextComma == std::string::npos || nextBracket < nextComma))
            break;
        if (nextComma != std::string::npos)
            pos = nextComma + 1;
        else
            break;
    }

    return results;
}
