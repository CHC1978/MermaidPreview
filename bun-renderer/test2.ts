import { JSDOM } from 'jsdom';

const dom = new JSDOM('<!DOCTYPE html><html><body><div id="container"></div></body></html>');

// Polyfill browser globals
Object.assign(globalThis, {
    document: dom.window.document,
    window: dom.window,
    navigator: dom.window.navigator,
    DOMParser: dom.window.DOMParser,
    XMLSerializer: dom.window.XMLSerializer,
    HTMLElement: dom.window.HTMLElement,
});

// Polyfill SVG methods that mermaid requires but jsdom doesn't provide
const SVGProto = (dom.window as any).SVGElement?.prototype || dom.window.HTMLElement.prototype;

// getBBox: estimate based on text content
SVGProto.getBBox = function(this: any) {
    const text = this.textContent || '';
    const fontSize = parseFloat(this.getAttribute?.('font-size') || '14');
    const charWidth = fontSize * 0.6;
    return {
        x: 0, y: 0,
        width: Math.max(text.length * charWidth, 20),
        height: fontSize * 1.4
    };
};

// getComputedTextLength
SVGProto.getComputedTextLength = function(this: any) {
    const text = this.textContent || '';
    const fontSize = parseFloat(this.getAttribute?.('font-size') || '14');
    return text.length * fontSize * 0.6;
};

// getBoundingClientRect
if (!SVGProto.getBoundingClientRect) {
    SVGProto.getBoundingClientRect = function(this: any) {
        const bbox = this.getBBox?.() || { x: 0, y: 0, width: 100, height: 20 };
        return { ...bbox, top: bbox.y, left: bbox.x, right: bbox.x + bbox.width, bottom: bbox.y + bbox.height };
    };
}

// Also polyfill on Element.prototype for non-SVG elements
const ElProto = dom.window.HTMLElement.prototype as any;
if (!ElProto.getBBox) {
    ElProto.getBBox = SVGProto.getBBox;
}

const mermaid = (await import('mermaid')).default;

mermaid.initialize({
    startOnLoad: false,
    securityLevel: 'strict',
    theme: 'default',
});

// Test multiple diagram types
const diagrams = [
    { name: 'flowchart', code: 'graph TD\n    A[Start] --> B[End]' },
    { name: 'sequence', code: 'sequenceDiagram\n    Alice->>Bob: Hello\n    Bob-->>Alice: Hi' },
    { name: 'pie', code: 'pie\n    "A" : 40\n    "B" : 60' },
];

for (const d of diagrams) {
    try {
        const { svg } = await mermaid.render(`test-${d.name}`, d.code);
        console.log(`${d.name}: SUCCESS (${svg.length} bytes)`);
    } catch (e: any) {
        console.log(`${d.name}: FAILED - ${e.message}`);
    }
}
