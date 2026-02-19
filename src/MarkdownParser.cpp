#include <windows.h>
#include "MarkdownParser.h"
#include "plugin.h"
#include <cwctype>
#include <algorithm>
#include <sstream>

// ============================================================================
// ExtractMermaidBlocks - unchanged from original
// ============================================================================
std::vector<MermaidBlock> MarkdownParser::ExtractMermaidBlocks(const std::wstring& content)
{
    std::vector<MermaidBlock> blocks;
    bool inBlock = false;
    MermaidBlock current;
    int lineNum = 0;

    size_t pos = 0;
    while (pos < content.size()) {
        size_t eol = content.find(L'\n', pos);
        if (eol == std::wstring::npos)
            eol = content.size();

        std::wstring line = content.substr(pos, eol - pos);
        if (!line.empty() && line.back() == L'\r')
            line.pop_back();

        if (!inBlock) {
            size_t i = 0;
            while (i < line.size() && (line[i] == L' ' || line[i] == L'\t'))
                i++;
            if (line.size() - i >= 10 &&
                line.substr(i, 10) == L"```mermaid") {
                size_t j = i + 10;
                while (j < line.size() && (line[j] == L' ' || line[j] == L'\t'))
                    j++;
                if (j >= line.size()) {
                    inBlock = true;
                    current = {};
                    current.startLine = lineNum;
                }
            }
        } else {
            size_t i = 0;
            while (i < line.size() && (line[i] == L' ' || line[i] == L'\t'))
                i++;
            if (line.size() - i >= 3 && line.substr(i, 3) == L"```") {
                size_t j = i + 3;
                while (j < line.size() && (line[j] == L' ' || line[j] == L'\t'))
                    j++;
                if (j >= line.size()) {
                    current.endLine = lineNum;
                    if (!current.code.empty() && current.code.back() == L'\n')
                        current.code.pop_back();
                    blocks.push_back(std::move(current));
                    inBlock = false;
                }
            } else {
                current.code += line;
                current.code += L'\n';
            }
        }

        lineNum++;
        pos = eol + 1;
    }

    return blocks;
}

// ============================================================================
// GetDocumentContent - unchanged from original
// ============================================================================
std::wstring MarkdownParser::GetDocumentContent(HWND hwndView)
{
    std::wstring content;

    UINT_PTR totalLines = (UINT_PTR)SendMessage(
        hwndView, EE_GET_LINES, (WPARAM)0, 0);

    if (totalLines == 0)
        return content;

    content.reserve(totalLines * 80);

    for (UINT_PTR i = 0; i < totalLines; i++) {
        GET_LINE_INFO gli = {};
        gli.cch = 0;
        gli.flags = 0;
        gli.yLine = i;

        UINT_PTR cch = (UINT_PTR)SendMessage(
            hwndView, EE_GET_LINEW, (WPARAM)&gli, (LPARAM)nullptr);

        if (cch == 0)
            continue;

        std::wstring lineBuf(cch, L'\0');
        gli.cch = cch;
        SendMessage(hwndView, EE_GET_LINEW, (WPARAM)&gli, (LPARAM)lineBuf.data());

        while (!lineBuf.empty() && lineBuf.back() == L'\0')
            lineBuf.pop_back();

        content += lineBuf;
        content += L'\n';
    }

    return content;
}

// ============================================================================
// HtmlEscape
// ============================================================================
std::wstring MarkdownParser::HtmlEscape(const std::wstring& text)
{
    std::wstring out;
    out.reserve(text.size() + text.size() / 8);
    for (wchar_t ch : text) {
        switch (ch) {
        case L'&':  out += L"&amp;";  break;
        case L'<':  out += L"&lt;";   break;
        case L'>':  out += L"&gt;";   break;
        case L'"':  out += L"&quot;"; break;
        default:    out += ch;        break;
        }
    }
    return out;
}

// ============================================================================
// UrlEncode - for mermaid data attributes
// ============================================================================
std::wstring MarkdownParser::UrlEncode(const std::wstring& text)
{
    std::wstring out;
    out.reserve(text.size() * 2);
    for (wchar_t ch : text) {
        if ((ch >= L'A' && ch <= L'Z') || (ch >= L'a' && ch <= L'z') ||
            (ch >= L'0' && ch <= L'9') || ch == L'-' || ch == L'_' ||
            ch == L'.' || ch == L'~' || ch == L' ') {
            if (ch == L' ')
                out += L"%20";
            else
                out += ch;
        } else if (ch == L'\n') {
            out += L"%0A";
        } else if (ch == L'\r') {
            out += L"%0D";
        } else if (ch < 0x80) {
            // Single-byte ASCII: %XX
            wchar_t buf[8];
            swprintf_s(buf, L"%%%02X", (unsigned)ch);
            out += buf;
        } else {
            // Multi-byte: convert to UTF-8 then percent-encode each byte
            char utf8[8] = {};
            int len = WideCharToMultiByte(CP_UTF8, 0, &ch, 1, utf8, sizeof(utf8), nullptr, nullptr);
            for (int i = 0; i < len; i++) {
                wchar_t buf[8];
                swprintf_s(buf, L"%%%02X", (unsigned char)utf8[i]);
                out += buf;
            }
        }
    }
    return out;
}

// ============================================================================
// IsSafeUrl - Block dangerous URL schemes (javascript:, data:, vbscript:)
// ============================================================================
bool MarkdownParser::IsSafeUrl(const std::wstring& url)
{
    // Trim leading whitespace
    size_t i = 0;
    while (i < url.size() && (url[i] == L' ' || url[i] == L'\t')) i++;
    if (i >= url.size()) return true; // empty is safe (renders nothing)

    // Lowercase the scheme prefix for case-insensitive comparison
    std::wstring lower;
    for (size_t j = i; j < url.size() && j < i + 16; j++)
        lower += (wchar_t)towlower(url[j]);

    if (lower.find(L"javascript:") == 0) return false;
    if (lower.find(L"vbscript:") == 0) return false;
    if (lower.find(L"data:") == 0) return false;

    return true;
}

// ============================================================================
// GenerateSlug - Convert heading text to URL-safe anchor id
// e.g. "My Heading!" -> "my-heading", "Hello World 123" -> "hello-world-123"
// ============================================================================
std::wstring MarkdownParser::GenerateSlug(const std::wstring& text)
{
    std::wstring slug;
    slug.reserve(text.size());
    bool prevDash = false;

    for (wchar_t ch : text) {
        if ((ch >= L'A' && ch <= L'Z')) {
            slug += (wchar_t)(ch + 32); // lowercase
            prevDash = false;
        } else if ((ch >= L'a' && ch <= L'z') || (ch >= L'0' && ch <= L'9')) {
            slug += ch;
            prevDash = false;
        } else if (ch > 0x7F) {
            // Keep non-ASCII characters (CJK, etc.) as-is
            slug += ch;
            prevDash = false;
        } else if (ch == L' ' || ch == L'-' || ch == L'_') {
            if (!slug.empty() && !prevDash) {
                slug += L'-';
                prevDash = true;
            }
        }
        // Other ASCII punctuation is stripped
    }

    // Trim trailing dash
    while (!slug.empty() && slug.back() == L'-')
        slug.pop_back();

    return slug;
}

// ============================================================================
// IsHorizontalRule
// ============================================================================
bool MarkdownParser::IsHorizontalRule(const std::wstring& line)
{
    size_t i = 0;
    while (i < line.size() && line[i] == L' ') i++;
    if (i >= line.size()) return false;

    wchar_t ch = line[i];
    if (ch != L'-' && ch != L'*' && ch != L'_') return false;

    int count = 0;
    while (i < line.size()) {
        if (line[i] == ch)
            count++;
        else if (line[i] != L' ')
            return false;
        i++;
    }
    return count >= 3;
}

// ============================================================================
// IsTableSeparator - |---|---|
// ============================================================================
bool MarkdownParser::IsTableSeparator(const std::wstring& line)
{
    size_t i = 0;
    while (i < line.size() && line[i] == L' ') i++;
    if (i >= line.size() || line[i] != L'|') return false;

    bool hasDash = false;
    for (size_t j = i; j < line.size(); j++) {
        if (line[j] == L'-') hasDash = true;
        else if (line[j] != L'|' && line[j] != L':' && line[j] != L' ')
            return false;
    }
    return hasDash;
}

// ============================================================================
// ParseTableRow
// ============================================================================
std::vector<std::wstring> MarkdownParser::ParseTableRow(const std::wstring& line)
{
    std::vector<std::wstring> cells;
    std::wstring trimmed = line;

    // Trim leading/trailing whitespace
    size_t s = 0, e = trimmed.size();
    while (s < e && trimmed[s] == L' ') s++;
    while (e > s && trimmed[e - 1] == L' ') e--;
    trimmed = trimmed.substr(s, e - s);

    // Remove leading/trailing pipes
    if (!trimmed.empty() && trimmed.front() == L'|')
        trimmed = trimmed.substr(1);
    if (!trimmed.empty() && trimmed.back() == L'|')
        trimmed.pop_back();

    // Split by |
    size_t pos = 0;
    while (pos < trimmed.size()) {
        size_t pipe = trimmed.find(L'|', pos);
        if (pipe == std::wstring::npos) pipe = trimmed.size();
        std::wstring cell = trimmed.substr(pos, pipe - pos);
        // Trim cell
        size_t cs = 0, ce = cell.size();
        while (cs < ce && cell[cs] == L' ') cs++;
        while (ce > cs && cell[ce - 1] == L' ') ce--;
        cells.push_back(cell.substr(cs, ce - cs));
        pos = pipe + 1;
    }
    return cells;
}

// ============================================================================
// ProcessInline - Handle inline formatting
// ============================================================================
std::wstring MarkdownParser::ProcessInline(const std::wstring& text)
{
    std::wstring out;
    out.reserve(text.size() + text.size() / 4);
    size_t len = text.size();
    size_t i = 0;

    while (i < len) {
        // --- Escaped character ---
        if (text[i] == L'\\' && i + 1 < len) {
            wchar_t next = text[i + 1];
            if (next == L'\\' || next == L'`' || next == L'*' || next == L'_' ||
                next == L'{' || next == L'}' || next == L'[' || next == L']' ||
                next == L'(' || next == L')' || next == L'#' || next == L'+' ||
                next == L'-' || next == L'.' || next == L'!' || next == L'|' ||
                next == L'~') {
                out += HtmlEscape(std::wstring(1, next));
                i += 2;
                continue;
            }
        }

        // --- Inline code: `code` ---
        if (text[i] == L'`') {
            // Count opening backticks
            size_t btStart = i;
            int btCount = 0;
            while (i < len && text[i] == L'`') { btCount++; i++; }

            // Find matching closing backticks
            size_t closePos = std::wstring::npos;
            for (size_t j = i; j <= len - btCount; j++) {
                bool match = true;
                for (int k = 0; k < btCount; k++) {
                    if (text[j + k] != L'`') { match = false; break; }
                }
                if (match) {
                    // Make sure it's exactly btCount backticks
                    if (j + btCount < len && text[j + btCount] == L'`')
                        continue;
                    closePos = j;
                    break;
                }
            }

            if (closePos != std::wstring::npos) {
                std::wstring code = text.substr(i, closePos - i);
                // Trim single leading/trailing space
                if (code.size() >= 2 && code.front() == L' ' && code.back() == L' ')
                    code = code.substr(1, code.size() - 2);
                out += L"<code>";
                out += HtmlEscape(code);
                out += L"</code>";
                i = closePos + btCount;
            } else {
                // No closing backticks: output literally
                for (int k = 0; k < btCount; k++) out += L'`';
            }
            continue;
        }

        // --- Image: ![alt](url) ---
        if (text[i] == L'!' && i + 1 < len && text[i + 1] == L'[') {
            size_t altStart = i + 2;
            size_t altEnd = text.find(L']', altStart);
            if (altEnd != std::wstring::npos && altEnd + 1 < len && text[altEnd + 1] == L'(') {
                size_t urlStart = altEnd + 2;
                size_t urlEnd = text.find(L')', urlStart);
                if (urlEnd != std::wstring::npos) {
                    std::wstring alt = text.substr(altStart, altEnd - altStart);
                    std::wstring url = text.substr(urlStart, urlEnd - urlStart);
                    if (IsSafeUrl(url)) {
                        out += L"<img src=\"";
                        out += HtmlEscape(url);
                        out += L"\" alt=\"";
                        out += HtmlEscape(alt);
                        out += L"\">";
                    } else {
                        out += L"[image blocked: unsafe URL]";
                    }
                    i = urlEnd + 1;
                    continue;
                }
            }
        }

        // --- Link: [text](url) ---
        if (text[i] == L'[') {
            size_t textEnd = std::wstring::npos;
            int depth = 1;
            for (size_t j = i + 1; j < len; j++) {
                if (text[j] == L'[') depth++;
                else if (text[j] == L']') { depth--; if (depth == 0) { textEnd = j; break; } }
            }
            if (textEnd != std::wstring::npos && textEnd + 1 < len && text[textEnd + 1] == L'(') {
                size_t urlStart = textEnd + 2;
                size_t urlEnd = text.find(L')', urlStart);
                if (urlEnd != std::wstring::npos) {
                    std::wstring linkText = text.substr(i + 1, textEnd - i - 1);
                    std::wstring url = text.substr(urlStart, urlEnd - urlStart);
                    if (IsSafeUrl(url)) {
                        out += L"<a href=\"";
                        out += HtmlEscape(url);
                        out += L"\">";
                        out += ProcessInline(linkText);
                        out += L"</a>";
                    } else {
                        out += ProcessInline(linkText);
                    }
                    i = urlEnd + 1;
                    continue;
                }
            }
        }

        // --- Bold + Italic: ***text*** ---
        if (i + 2 < len && text[i] == L'*' && text[i + 1] == L'*' && text[i + 2] == L'*') {
            size_t end = text.find(L"***", i + 3);
            if (end != std::wstring::npos) {
                out += L"<strong><em>";
                out += ProcessInline(text.substr(i + 3, end - i - 3));
                out += L"</em></strong>";
                i = end + 3;
                continue;
            }
        }

        // --- Bold: **text** ---
        if (i + 1 < len && text[i] == L'*' && text[i + 1] == L'*') {
            size_t end = text.find(L"**", i + 2);
            if (end != std::wstring::npos) {
                out += L"<strong>";
                out += ProcessInline(text.substr(i + 2, end - i - 2));
                out += L"</strong>";
                i = end + 2;
                continue;
            }
        }

        // --- Italic: *text* ---
        if (text[i] == L'*' && i + 1 < len && text[i + 1] != L' ') {
            size_t end = text.find(L'*', i + 1);
            if (end != std::wstring::npos && text[end - 1] != L' ') {
                out += L"<em>";
                out += ProcessInline(text.substr(i + 1, end - i - 1));
                out += L"</em>";
                i = end + 1;
                continue;
            }
        }

        // --- Strikethrough: ~~text~~ ---
        if (i + 1 < len && text[i] == L'~' && text[i + 1] == L'~') {
            size_t end = text.find(L"~~", i + 2);
            if (end != std::wstring::npos) {
                out += L"<del>";
                out += ProcessInline(text.substr(i + 2, end - i - 2));
                out += L"</del>";
                i = end + 2;
                continue;
            }
        }

        // --- HTML entities ---
        if (text[i] == L'&') { out += L"&amp;"; i++; continue; }
        if (text[i] == L'<') { out += L"&lt;";  i++; continue; }
        if (text[i] == L'>') { out += L"&gt;";  i++; continue; }

        // --- Normal character ---
        out += text[i];
        i++;
    }
    return out;
}

// ============================================================================
// Helper: split content into lines
// ============================================================================
static std::vector<std::wstring> SplitLines(const std::wstring& content)
{
    std::vector<std::wstring> lines;
    size_t pos = 0;
    while (pos < content.size()) {
        size_t eol = content.find(L'\n', pos);
        if (eol == std::wstring::npos) eol = content.size();
        std::wstring line = content.substr(pos, eol - pos);
        if (!line.empty() && line.back() == L'\r')
            line.pop_back();
        lines.push_back(std::move(line));
        pos = eol + 1;
    }
    return lines;
}

// ============================================================================
// Helper: trim leading whitespace, return indent count
// ============================================================================
static std::wstring TrimLeft(const std::wstring& s, size_t* indentOut = nullptr)
{
    size_t i = 0;
    while (i < s.size() && (s[i] == L' ' || s[i] == L'\t')) i++;
    if (indentOut) *indentOut = i;
    return s.substr(i);
}

// ============================================================================
// Helper: check if line is blank
// ============================================================================
static bool IsBlank(const std::wstring& line)
{
    for (wchar_t ch : line)
        if (ch != L' ' && ch != L'\t') return false;
    return true;
}

// ============================================================================
// ConvertToHtml - Main markdown-to-HTML converter
// ============================================================================
std::wstring MarkdownParser::ConvertToHtml(const std::wstring& markdown)
{
    auto lines = SplitLines(markdown);
    std::wstring html;
    html.reserve(markdown.size() * 2);

    size_t n = lines.size();
    size_t i = 0;
    int mermaidIdx = 0;

    // Track if we're accumulating a paragraph
    std::wstring paraAccum;
    int paraStartLine = -1;
    int paraEndLine = -1;

    auto flushParagraph = [&]() {
        if (!paraAccum.empty()) {
            html += L"<p data-line-start=\"" + std::to_wstring(paraStartLine)
                 + L"\" data-line-end=\"" + std::to_wstring(paraEndLine) + L"\">";
            html += ProcessInline(paraAccum);
            html += L"</p>\n";
            paraAccum.clear();
            paraStartLine = -1;
            paraEndLine = -1;
        }
    };

    while (i < n) {
        const std::wstring& rawLine = lines[i];
        std::wstring trimmed = TrimLeft(rawLine);

        // --- Blank line: flush paragraph ---
        if (IsBlank(rawLine)) {
            flushParagraph();
            i++;
            continue;
        }

        // --- HTML anchor tag: <a id="..."></a> or <a name="..."></a> ---
        // Pass through as raw HTML so internal links (#anchor) work
        if (trimmed.size() >= 8 && trimmed[0] == L'<' && trimmed[1] == L'a' && trimmed[2] == L' ') {
            // Check for <a id="..."></a> or <a name="..."></a> pattern
            size_t closeTag = trimmed.find(L"</a>");
            if (closeTag != std::wstring::npos) {
                size_t idPos = trimmed.find(L"id=\"");
                size_t namePos = trimmed.find(L"name=\"");
                if (idPos != std::wstring::npos || namePos != std::wstring::npos) {
                    flushParagraph();
                    // Output the anchor tag as-is (raw HTML passthrough)
                    html += trimmed;
                    html += L"\n";
                    i++;
                    continue;
                }
            }
        }

        // --- Fenced code block: ``` or ~~~ ---
        if (trimmed.size() >= 3 &&
            ((trimmed[0] == L'`' && trimmed[1] == L'`' && trimmed[2] == L'`') ||
             (trimmed[0] == L'~' && trimmed[1] == L'~' && trimmed[2] == L'~'))) {
            flushParagraph();
            wchar_t fence = trimmed[0];

            // Count fence chars
            size_t fenceLen = 0;
            while (fenceLen < trimmed.size() && trimmed[fenceLen] == fence) fenceLen++;

            // Extract language tag
            std::wstring lang = trimmed.substr(fenceLen);
            // Trim whitespace from lang
            size_t ls = 0;
            while (ls < lang.size() && lang[ls] == L' ') ls++;
            size_t le = lang.size();
            while (le > ls && lang[le - 1] == L' ') le--;
            lang = lang.substr(ls, le - ls);

            // Lowercase lang for comparison
            std::wstring langLower = lang;
            for (auto& c : langLower) c = towlower(c);

            // Collect code block content
            int codeBlockStartLine = (int)i; // opening fence line
            std::wstring codeContent;
            i++;
            while (i < n) {
                std::wstring ct = TrimLeft(lines[i]);
                // Check for closing fence
                size_t closeFenceLen = 0;
                while (closeFenceLen < ct.size() && ct[closeFenceLen] == fence) closeFenceLen++;
                if (closeFenceLen >= fenceLen) {
                    std::wstring rest = ct.substr(closeFenceLen);
                    bool onlyWhitespace = true;
                    for (wchar_t ch : rest)
                        if (ch != L' ' && ch != L'\t') { onlyWhitespace = false; break; }
                    if (onlyWhitespace) {
                        i++;
                        break;
                    }
                }
                codeContent += lines[i];
                codeContent += L'\n';
                i++;
            }
            int codeBlockEndLine = (int)(i - 1); // closing fence line
            // Remove trailing newline
            if (!codeContent.empty() && codeContent.back() == L'\n')
                codeContent.pop_back();

            // Mermaid block → placeholder div
            if (langLower == L"mermaid") {
                std::wstring id = L"mermaid-placeholder-" + std::to_wstring(mermaidIdx++);
                html += L"<div class=\"mermaid-container\" data-mermaid-id=\"";
                html += id;
                html += L"\" data-mermaid-src=\"";
                html += UrlEncode(codeContent);
                html += L"\" data-line-start=\"";
                html += std::to_wstring(codeBlockStartLine);
                html += L"\" data-line-end=\"";
                html += std::to_wstring(codeBlockEndLine);
                html += L"\"></div>\n";
            } else {
                // Regular code block
                html += L"<pre><code";
                if (!lang.empty()) {
                    html += L" class=\"language-";
                    html += HtmlEscape(lang);
                    html += L"\"";
                }
                html += L">";
                html += HtmlEscape(codeContent);
                html += L"</code></pre>\n";
            }
            continue;
        }

        // --- ATX Heading: # through ###### ---
        if (trimmed.size() >= 2 && trimmed[0] == L'#') {
            int level = 0;
            size_t hi = 0;
            while (hi < trimmed.size() && hi < 6 && trimmed[hi] == L'#') { level++; hi++; }
            if (hi < trimmed.size() && trimmed[hi] == L' ') {
                flushParagraph();
                std::wstring headText = trimmed.substr(hi + 1);
                // Remove trailing #s
                size_t te = headText.size();
                while (te > 0 && headText[te - 1] == L'#') te--;
                while (te > 0 && headText[te - 1] == L' ') te--;
                headText = headText.substr(0, te);

                std::wstring slug = GenerateSlug(headText);
                html += L"<h" + std::to_wstring(level);
                if (!slug.empty()) {
                    html += L" id=\"";
                    html += HtmlEscape(slug);
                    html += L"\"";
                }
                html += L" data-line-start=\""
                     + std::to_wstring((int)i) + L"\" data-line-end=\""
                     + std::to_wstring((int)i) + L"\">";
                html += ProcessInline(headText);
                html += L"</h" + std::to_wstring(level) + L">\n";
                i++;
                continue;
            }
        }

        // --- Horizontal rule ---
        if (IsHorizontalRule(trimmed)) {
            flushParagraph();
            html += L"<hr>\n";
            i++;
            continue;
        }

        // --- Blockquote: > ---
        if (trimmed.size() >= 1 && trimmed[0] == L'>') {
            flushParagraph();
            std::wstring bqContent;
            while (i < n) {
                std::wstring t = TrimLeft(lines[i]);
                if (t.empty() || t[0] != L'>') break;
                // Remove > and optional space
                std::wstring bqLine = t.substr(1);
                if (!bqLine.empty() && bqLine[0] == L' ')
                    bqLine = bqLine.substr(1);
                bqContent += bqLine + L'\n';
                i++;
            }
            // Recursively convert blockquote content
            html += L"<blockquote>\n";
            html += ConvertToHtml(bqContent);
            html += L"</blockquote>\n";
            continue;
        }

        // --- Table: | col | col | ---
        if (trimmed.size() >= 1 && trimmed[0] == L'|' &&
            i + 1 < n && IsTableSeparator(TrimLeft(lines[i + 1]))) {
            flushParagraph();
            // Header row
            auto headerCells = ParseTableRow(trimmed);
            i++; // skip separator
            i++;

            html += L"<table>\n<thead>\n<tr>\n";
            for (auto& cell : headerCells) {
                html += L"<th>";
                html += ProcessInline(cell);
                html += L"</th>\n";
            }
            html += L"</tr>\n</thead>\n<tbody>\n";

            // Body rows
            while (i < n) {
                std::wstring t = TrimLeft(lines[i]);
                if (t.empty() || t[0] != L'|') break;
                auto cells = ParseTableRow(t);
                html += L"<tr>\n";
                for (size_t ci = 0; ci < cells.size(); ci++) {
                    html += L"<td>";
                    html += ProcessInline(cells[ci]);
                    html += L"</td>\n";
                }
                html += L"</tr>\n";
                i++;
            }
            html += L"</tbody>\n</table>\n";
            continue;
        }

        // --- Unordered list: - / * / + ---
        if (trimmed.size() >= 2 &&
            (trimmed[0] == L'-' || trimmed[0] == L'*' || trimmed[0] == L'+') &&
            trimmed[1] == L' ') {
            flushParagraph();
            html += L"<ul>\n";
            while (i < n) {
                std::wstring t = TrimLeft(lines[i]);
                if (t.size() < 2) break;

                // Check for list item or task list
                bool isItem = (t[0] == L'-' || t[0] == L'*' || t[0] == L'+') && t[1] == L' ';
                if (!isItem) break;

                int itemStartLine = (int)i;
                std::wstring itemText = t.substr(2);

                // Check for task list: [ ] or [x]
                bool isTask = false;
                bool isChecked = false;
                if (itemText.size() >= 3 && itemText[0] == L'[') {
                    if (itemText[1] == L' ' && itemText[2] == L']') {
                        isTask = true; isChecked = false;
                        itemText = itemText.substr(3);
                        if (!itemText.empty() && itemText[0] == L' ')
                            itemText = itemText.substr(1);
                    } else if ((itemText[1] == L'x' || itemText[1] == L'X') && itemText[2] == L']') {
                        isTask = true; isChecked = true;
                        itemText = itemText.substr(3);
                        if (!itemText.empty() && itemText[0] == L' ')
                            itemText = itemText.substr(1);
                    }
                }

                // Collect continuation lines (indented)
                i++;
                while (i < n && !IsBlank(lines[i])) {
                    size_t indent = 0;
                    std::wstring ct = TrimLeft(lines[i], &indent);
                    if (indent < 2) break;
                    // Check if next line is a new list item
                    if (ct.size() >= 2 &&
                        (ct[0] == L'-' || ct[0] == L'*' || ct[0] == L'+') &&
                        ct[1] == L' ')
                        break;
                    itemText += L" " + ct;
                    i++;
                }

                int itemEndLine = (int)(i - 1);
                std::wstring lineAttr = L" data-line-start=\"" + std::to_wstring(itemStartLine)
                                      + L"\" data-line-end=\"" + std::to_wstring(itemEndLine) + L"\"";
                if (isTask) {
                    html += L"<li class=\"task-list-item\"" + lineAttr + L"><input type=\"checkbox\" disabled";
                    if (isChecked) html += L" checked";
                    html += L"> ";
                    html += ProcessInline(itemText);
                    html += L"</li>\n";
                } else {
                    html += L"<li" + lineAttr + L">";
                    html += ProcessInline(itemText);
                    html += L"</li>\n";
                }
            }
            html += L"</ul>\n";
            continue;
        }

        // --- Ordered list: 1. 2. etc ---
        {
            size_t di = 0;
            while (di < trimmed.size() && trimmed[di] >= L'0' && trimmed[di] <= L'9') di++;
            if (di > 0 && di < trimmed.size() && trimmed[di] == L'.' &&
                di + 1 < trimmed.size() && trimmed[di + 1] == L' ') {
                flushParagraph();
                html += L"<ol>\n";
                while (i < n) {
                    std::wstring t = TrimLeft(lines[i]);
                    size_t d = 0;
                    while (d < t.size() && t[d] >= L'0' && t[d] <= L'9') d++;
                    if (d == 0 || d >= t.size() || t[d] != L'.' ||
                        d + 1 >= t.size() || t[d + 1] != L' ')
                        break;

                    int olItemStart = (int)i;
                    std::wstring itemText = t.substr(d + 2);

                    // Collect continuation lines
                    i++;
                    while (i < n && !IsBlank(lines[i])) {
                        size_t indent = 0;
                        std::wstring ct = TrimLeft(lines[i], &indent);
                        if (indent < 2) break;
                        // Check if next line is a new list item
                        size_t nd = 0;
                        while (nd < ct.size() && ct[nd] >= L'0' && ct[nd] <= L'9') nd++;
                        if (nd > 0 && nd < ct.size() && ct[nd] == L'.')
                            break;
                        itemText += L" " + ct;
                        i++;
                    }

                    int olItemEnd = (int)(i - 1);
                    html += L"<li data-line-start=\"" + std::to_wstring(olItemStart)
                         + L"\" data-line-end=\"" + std::to_wstring(olItemEnd) + L"\">";
                    html += ProcessInline(itemText);
                    html += L"</li>\n";
                }
                html += L"</ol>\n";
                continue;
            }
        }

        // --- Setext heading (underline-style): ==== or ---- ---
        if (i + 1 < n) {
            std::wstring nextTrimmed = TrimLeft(lines[i + 1]);
            bool isH1 = !nextTrimmed.empty() && nextTrimmed.find_first_not_of(L"= ") == std::wstring::npos
                         && nextTrimmed.find(L'=') != std::wstring::npos;
            bool isH2 = !nextTrimmed.empty() && nextTrimmed.find_first_not_of(L"- ") == std::wstring::npos
                         && nextTrimmed.find(L'-') != std::wstring::npos
                         && std::count(nextTrimmed.begin(), nextTrimmed.end(), L'-') >= 3;
            if (isH1 || isH2) {
                flushParagraph();
                int level = isH1 ? 1 : 2;
                std::wstring slug = GenerateSlug(trimmed);
                html += L"<h" + std::to_wstring(level);
                if (!slug.empty()) {
                    html += L" id=\"";
                    html += HtmlEscape(slug);
                    html += L"\"";
                }
                html += L" data-line-start=\""
                     + std::to_wstring((int)i) + L"\" data-line-end=\""
                     + std::to_wstring((int)(i + 1)) + L"\">";
                html += ProcessInline(trimmed);
                html += L"</h" + std::to_wstring(level) + L">\n";
                i += 2;
                continue;
            }
        }

        // --- Paragraph text (accumulate) ---
        if (paraStartLine < 0) paraStartLine = (int)i;
        paraEndLine = (int)i;
        if (!paraAccum.empty())
            paraAccum += L" ";
        paraAccum += trimmed;
        i++;
    }

    // Flush any remaining paragraph
    flushParagraph();

    return html;
}
