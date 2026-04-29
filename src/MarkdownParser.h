#pragma once

#include <windows.h>
#include <string>
#include <vector>

struct MermaidBlock {
    std::wstring code;
    int startLine;
    int endLine;
};

// One node-label range inside a flowchart mermaid block source. Produced
// by ExtractFlowchartNodes for the inline-edit pipeline (M2): JS sends
// {blockId, nodeId, newLabel} back, C++ resolves the matching ref and
// performs the in-place substring replacement.
struct MermaidNodeRef {
    std::wstring nodeId;          // e.g. "A"
    int          lineOffsetInBlock = 0; // 0-based, relative to first line of block source
    size_t       labelStart = 0;  // byte offset within that line (start of label content)
    size_t       labelEnd   = 0;  // exclusive end
    wchar_t      openBracket = 0; // '[' '{' '(' — used to pick the matching close + decide quoting
    bool         isQuoted = false; // true → label was already wrapped in "..."
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

    // Index every flowchart node label in a mermaid block source. Supports
    // square `A[label]`, round `A(label)`, decision `A{label}`, and the
    // quoted variants `A["..."]` etc. Skips %%directives, comment lines,
    // and `subgraph`/`end`. Returns at most one ref per nodeId (the first
    // *defining* occurrence; later references like `A --> B` are ignored).
    static std::vector<MermaidNodeRef>
    ExtractFlowchartNodes(const std::wstring& blockSource);

    // HTML-escape special characters (public for use by other modules)
    static std::wstring HtmlEscape(const std::wstring& text);

private:
    // Inline formatting: bold, italic, code, links, images, strikethrough.
    // `depth` guards against pathologically nested markdown (e.g.
    // `**[***x***](u)**`) overflowing the call stack.
    static std::wstring ProcessInline(const std::wstring& text, int depth = 0);

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
