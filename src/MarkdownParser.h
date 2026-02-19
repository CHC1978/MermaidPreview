#pragma once

#include <windows.h>
#include <string>
#include <vector>

struct MermaidBlock {
    std::wstring code;
    int startLine;
    int endLine;
};

class MarkdownParser {
public:
    // Extract all ```mermaid code blocks from document content
    static std::vector<MermaidBlock> ExtractMermaidBlocks(const std::wstring& content);

    // Get full document content from EmEditor view window
    static std::wstring GetDocumentContent(HWND hwndView);

    // Convert raw Markdown to HTML (C++ native, no JS dependency)
    // Mermaid blocks become <div class="mermaid-container" data-mermaid-src="...">
    static std::wstring ConvertToHtml(const std::wstring& markdown);

private:
    // Inline formatting: bold, italic, code, links, images, strikethrough
    static std::wstring ProcessInline(const std::wstring& text);

    // HTML-escape special characters
    static std::wstring HtmlEscape(const std::wstring& text);

    // URL-encode for data attributes
    static std::wstring UrlEncode(const std::wstring& text);

    // Check if a line is a horizontal rule (---, ***, ___)
    static bool IsHorizontalRule(const std::wstring& line);

    // Check if a line is a table separator (|---|---|)
    static bool IsTableSeparator(const std::wstring& line);

    // Parse a table row into cells
    static std::vector<std::wstring> ParseTableRow(const std::wstring& line);

    // Check if URL scheme is safe (block javascript:, data:, vbscript:)
    static bool IsSafeUrl(const std::wstring& url);

    // Generate a URL-safe slug from heading text (for id attributes)
    static std::wstring GenerateSlug(const std::wstring& text);
};
