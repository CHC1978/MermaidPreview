import { JSDOM } from 'jsdom';

// Set up DOM environment BEFORE importing mermaid
const dom = new JSDOM('<!DOCTYPE html><html><body><div id="container"></div></body></html>');

// Polyfill browser globals that mermaid needs
Object.assign(globalThis, {
    document: dom.window.document,
    window: dom.window,
    navigator: dom.window.navigator,
    DOMParser: dom.window.DOMParser,
    XMLSerializer: dom.window.XMLSerializer,
    HTMLElement: dom.window.HTMLElement,
    SVGElement: (dom.window as any).SVGElement || dom.window.HTMLElement,
});

// Now import mermaid (after DOM globals are set)
const mermaid = (await import('mermaid')).default;

mermaid.initialize({
    startOnLoad: false,
    securityLevel: 'strict',
    theme: 'default',
});

try {
    const { svg } = await mermaid.render('test-id', 'graph TD\n    A[Start] --> B[End]');
    console.log('SUCCESS! SVG length:', svg.length);
    console.log('First 300 chars:', svg.substring(0, 300));
} catch (e: any) {
    console.error('FAILED:', e.message || e);
    if (e.stack) console.error(e.stack);
}
