/**
 * Mermaid Renderer - Persistent Bun process
 * Communicates via stdin/stdout JSON (one JSON per line)
 *
 * Request:  {"type":"render","blocks":[{"id":"mmd-0","code":"graph TD\nA-->B"}],"theme":"default"}
 * Response: {"type":"result","results":[{"id":"mmd-0","svg":"<svg>...</svg>","error":null}]}
 *
 * Request:  {"type":"ping"}
 * Response: {"type":"pong"}
 */

import { JSDOM } from 'jsdom';

// ── Set up virtual DOM environment ──────────────────────────────────
const dom = new JSDOM('<!DOCTYPE html><html><body><div id="container"></div></body></html>');

Object.assign(globalThis, {
    document: dom.window.document,
    window: dom.window,
    navigator: dom.window.navigator,
    DOMParser: dom.window.DOMParser,
    XMLSerializer: dom.window.XMLSerializer,
    HTMLElement: dom.window.HTMLElement,
});

// ── Improved SVG measurement polyfill ───────────────────────────────
function computeBBox(el: any): { x: number; y: number; width: number; height: number } {
    const tag = el.tagName?.toLowerCase();

    if (tag === 'rect') {
        return {
            x: parseFloat(el.getAttribute('x') || '0'),
            y: parseFloat(el.getAttribute('y') || '0'),
            width: parseFloat(el.getAttribute('width') || '0'),
            height: parseFloat(el.getAttribute('height') || '0'),
        };
    }

    if (tag === 'circle') {
        const cx = parseFloat(el.getAttribute('cx') || '0');
        const cy = parseFloat(el.getAttribute('cy') || '0');
        const r = parseFloat(el.getAttribute('r') || '0');
        return { x: cx - r, y: cy - r, width: r * 2, height: r * 2 };
    }

    if (tag === 'line') {
        const x1 = parseFloat(el.getAttribute('x1') || '0');
        const y1 = parseFloat(el.getAttribute('y1') || '0');
        const x2 = parseFloat(el.getAttribute('x2') || '0');
        const y2 = parseFloat(el.getAttribute('y2') || '0');
        return {
            x: Math.min(x1, x2), y: Math.min(y1, y2),
            width: Math.abs(x2 - x1), height: Math.abs(y2 - y1),
        };
    }

    if (tag === 'text' || tag === 'tspan') {
        const text = el.textContent || '';
        const fontSize = parseFloat(el.getAttribute('font-size') || '16');
        const w = text.length * fontSize * 0.6;
        return { x: 0, y: -fontSize, width: Math.max(w, 10), height: fontSize * 1.2 };
    }

    if (tag === 'foreignobject') {
        return {
            x: parseFloat(el.getAttribute('x') || '0'),
            y: parseFloat(el.getAttribute('y') || '0'),
            width: parseFloat(el.getAttribute('width') || '100'),
            height: parseFloat(el.getAttribute('height') || '20'),
        };
    }

    // Group elements: compute union of child bounding boxes
    if (tag === 'g' || tag === 'svg') {
        let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
        let found = false;

        const children = el.children || el.childNodes || [];
        for (let i = 0; i < children.length; i++) {
            const child = children[i];
            if (!child.tagName) continue;
            const ct = child.tagName.toLowerCase();
            if (ct === 'style' || ct === 'defs' || ct === 'marker') continue;

            let childBox: { x: number; y: number; width: number; height: number };
            try {
                childBox = child.getBBox ? child.getBBox() : computeBBox(child);
            } catch { continue; }

            if (childBox.width <= 0 && childBox.height <= 0) continue;

            // Apply transform offset
            const transform = child.getAttribute?.('transform') || '';
            const tm = transform.match(/translate\(\s*([\d.eE+-]+)\s*[,\s]\s*([\d.eE+-]+)\s*\)/);
            const tx = tm ? parseFloat(tm[1]) : 0;
            const ty = tm ? parseFloat(tm[2]) : 0;

            minX = Math.min(minX, childBox.x + tx);
            minY = Math.min(minY, childBox.y + ty);
            maxX = Math.max(maxX, childBox.x + tx + childBox.width);
            maxY = Math.max(maxY, childBox.y + ty + childBox.height);
            found = true;
        }

        if (found && minX !== Infinity) {
            return { x: minX, y: minY, width: maxX - minX, height: maxY - minY };
        }
    }

    // Fallback: estimate from text content
    const text = el.textContent || '';
    return { x: 0, y: 0, width: Math.max(text.length * 9, 20), height: 20 };
}

const SVGProto = (dom.window as any).SVGElement?.prototype || dom.window.HTMLElement.prototype;
SVGProto.getBBox = function (this: any) { return computeBBox(this); };

SVGProto.getComputedTextLength = function (this: any) {
    const text = this.textContent || '';
    const fontSize = parseFloat(this.getAttribute?.('font-size') || '16');
    return text.length * fontSize * 0.6;
};

if (!SVGProto.getBoundingClientRect) {
    SVGProto.getBoundingClientRect = function (this: any) {
        const b = computeBBox(this);
        return {
            x: b.x, y: b.y, width: b.width, height: b.height,
            top: b.y, left: b.x, right: b.x + b.width, bottom: b.y + b.height,
        };
    };
}

const ElProto = dom.window.HTMLElement.prototype as any;
if (!ElProto.getBBox) ElProto.getBBox = function (this: any) { return computeBBox(this); };

// ── SVG Post-processing: fix viewBox from actual element positions ───
function fixSvgBounds(svg: string): string {
    // Find all translate(x, y) on node groups
    const transforms = [...svg.matchAll(/transform="translate\(\s*([\d.eE+-]+)\s*[,\s]\s*([\d.eE+-]+)\s*\)"/g)];
    if (transforms.length === 0) return svg;

    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
    for (const m of transforms) {
        const x = parseFloat(m[1]);
        const y = parseFloat(m[2]);
        if (isNaN(x) || isNaN(y)) continue;
        minX = Math.min(minX, x);
        minY = Math.min(minY, y);
        maxX = Math.max(maxX, x);
        maxY = Math.max(maxY, y);
    }

    if (minX === Infinity) return svg;

    // Find max rect sizes for node padding
    let maxNodeW = 80, maxNodeH = 40;
    const rectMatches = [...svg.matchAll(/(?:width|height)="(\d+)"/g)];
    for (const r of rectMatches) {
        const val = parseInt(r[1]);
        if (val > 10 && val < 800) {
            if (r[0].startsWith('width')) maxNodeW = Math.max(maxNodeW, val);
            else maxNodeH = Math.max(maxNodeH, val);
        }
    }

    const padX = maxNodeW / 2 + 30;
    const padY = maxNodeH / 2 + 30;
    const vbX = minX - padX;
    const vbY = minY - padY;
    const width = maxX - minX + padX * 2;
    const height = maxY - minY + padY * 2;

    svg = svg.replace(/viewBox="[^"]*"/, `viewBox="${vbX} ${vbY} ${width} ${height}"`);
    svg = svg.replace(/max-width:\s*[\d.]+px;?/g, `max-width: ${Math.max(Math.ceil(width) + 50, 300)}px;`);

    return svg;
}

// ── Import and initialize mermaid ───────────────────────────────────
const mermaid = (await import('mermaid')).default;

let currentTheme = 'default';
mermaid.initialize({
    startOnLoad: false,
    securityLevel: 'strict',
    theme: currentTheme,
    flowchart: { useMaxWidth: true },
    sequence: { useMaxWidth: true },
});

// ── Signal readiness ────────────────────────────────────────────────
console.log(JSON.stringify({ type: 'ready' }));

// ── Process stdin line by line ──────────────────────────────────────
const decoder = new TextDecoder();
let buffer = '';

for await (const chunk of Bun.stdin.stream()) {
    buffer += decoder.decode(chunk, { stream: true });

    let newlineIdx: number;
    while ((newlineIdx = buffer.indexOf('\n')) !== -1) {
        const line = buffer.substring(0, newlineIdx).trim();
        buffer = buffer.substring(newlineIdx + 1);

        if (!line) continue;

        try {
            const req = JSON.parse(line);

            if (req.type === 'ping') {
                console.log(JSON.stringify({ type: 'pong' }));
                continue;
            }

            if (req.type === 'render') {
                // Validate theme (whitelist only)
                const VALID_THEMES = ['default', 'dark', 'neutral', 'forest', 'base'];
                const theme = (typeof req.theme === 'string' && VALID_THEMES.includes(req.theme))
                    ? req.theme : 'default';

                if (theme !== currentTheme) {
                    currentTheme = theme;
                    mermaid.initialize({
                        startOnLoad: false,
                        securityLevel: 'strict',
                        theme: currentTheme,
                        flowchart: { useMaxWidth: true },
                        sequence: { useMaxWidth: true },
                    });
                }

                const results: Array<{ id: string; svg: string | null; error: string | null }> = [];
                const blocks = Array.isArray(req.blocks) ? req.blocks : [];
                const MAX_CODE_LENGTH = 100000; // 100KB per block

                for (const block of blocks) {
                    // Validate block.id: must be string, alphanumeric + dash/underscore
                    if (typeof block.id !== 'string' || !/^[a-zA-Z0-9_-]+$/.test(block.id)) {
                        results.push({ id: String(block.id || 'invalid'), svg: null, error: 'Invalid block id' });
                        continue;
                    }
                    // Validate block.code: must be string with length limit
                    if (typeof block.code !== 'string') {
                        results.push({ id: block.id, svg: null, error: 'Invalid block code type' });
                        continue;
                    }
                    if (block.code.length > MAX_CODE_LENGTH) {
                        results.push({ id: block.id, svg: null, error: 'Code too long' });
                        continue;
                    }

                    try {
                        dom.window.document.body.innerHTML = '<div id="container"></div>';
                        const { svg } = await mermaid.render(block.id, block.code);
                        results.push({ id: block.id, svg: fixSvgBounds(svg), error: null });
                    } catch (e: any) {
                        results.push({
                            id: block.id,
                            svg: null,
                            error: (e.message || String(e)).substring(0, 500),
                        });
                    }
                }

                console.log(JSON.stringify({ type: 'result', results }));
                continue;
            }

            console.log(JSON.stringify({ type: 'error', message: 'Unknown request type: ' + req.type }));
        } catch (e: any) {
            console.log(JSON.stringify({ type: 'error', message: 'JSON parse error: ' + e.message }));
        }
    }
}
