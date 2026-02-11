import { JSDOM } from 'jsdom';
import { writeFileSync } from 'fs';

const dom = new JSDOM('<!DOCTYPE html><html><body><div id="container"></div></body></html>');
Object.assign(globalThis, {
    document: dom.window.document,
    window: dom.window,
    navigator: dom.window.navigator,
    DOMParser: dom.window.DOMParser,
    XMLSerializer: dom.window.XMLSerializer,
    HTMLElement: dom.window.HTMLElement,
});

const SVGProto = (dom.window as any).SVGElement?.prototype || dom.window.HTMLElement.prototype;
SVGProto.getBBox = function(this: any) {
    const text = this.textContent || '';
    const fontSize = parseFloat(this.getAttribute?.('font-size') || '14');
    return { x: 0, y: 0, width: Math.max(text.length * fontSize * 0.6, 20), height: fontSize * 1.4 };
};
SVGProto.getComputedTextLength = function(this: any) {
    const text = this.textContent || '';
    return text.length * parseFloat(this.getAttribute?.('font-size') || '14') * 0.6;
};
if (!SVGProto.getBoundingClientRect) {
    SVGProto.getBoundingClientRect = function(this: any) {
        const b = this.getBBox?.() || { x: 0, y: 0, width: 100, height: 20 };
        return { ...b, top: b.y, left: b.x, right: b.x + b.width, bottom: b.y + b.height };
    };
}
const ElProto = dom.window.HTMLElement.prototype as any;
if (!ElProto.getBBox) ElProto.getBBox = SVGProto.getBBox;

const mermaid = (await import('mermaid')).default;
mermaid.initialize({ startOnLoad: false, securityLevel: 'strict', theme: 'default' });

const code = `graph TD
    A[Start] --> B{Decision}
    B -->|Yes| C[Process A]
    B -->|No| D[Process B]
    C --> E[End]
    D --> E`;

const { svg } = await mermaid.render('test-flow', code);

// Write HTML for visual inspection
const html = `<!DOCTYPE html><html><body style="display:flex;justify-content:center;padding:40px;background:#f0f0f0">${svg}</body></html>`;
writeFileSync('test-output.html', html);
console.log('Wrote test-output.html');
console.log('SVG size:', svg.length, 'bytes');
