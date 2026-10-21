import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

/**
 * Parses an XML string into a DOM Document.
 */
export function parseXml(xmlString: string): Document | null {
  if (!xmlString) return null;
  return new DOMParser().parseFromString(xmlString, 'text/xml');
}

/**
 * Serializes a DOM Document or Element into an XML string.
 */
export function serializeXml(node: Node | null): string {
  if (!node) return '';
  return new XMLSerializer().serializeToString(node);
}

/**
 * Gets direct child elements matching a local tag name.
 */
export function getDirectChildren(parent: Node, localName?: string): Element[] {
  const results: Element[] = [];
  if (!parent || !parent.childNodes) return results;

  for (let i = 0; i < parent.childNodes.length; i++) {
    const child = parent.childNodes[i];
    if (child.nodeType === 1) { // Node.ELEMENT_NODE
      const el = child as Element;
      if (!localName || el.localName === localName || el.nodeName === localName || el.nodeName.endsWith(':' + localName)) {
        results.push(el);
      }
    }
  }
  return results;
}

/**
 * Recursively visits all Element nodes in a DOM subtree.
 */
export function walkElements(node: Node | null, callback: (el: Element) => void): void {
  if (!node) return;
  if (node.nodeType === 1) {
    callback(node as Element);
  }
  if (node.childNodes) {
    for (let i = 0; i < node.childNodes.length; i++) {
      walkElements(node.childNodes[i], callback);
    }
  }
}

/**
 * Extracts filename from path without Node 'path' module (browser compatible).
 */
export function getBasename(filePath: string): string {
  if (!filePath) return '';
  const parts = filePath.split(/[/\\]/);
  return parts[parts.length - 1] || '';
}

/**
 * Extracts file extension including dot without Node 'path' module (browser compatible).
 */
export function getExtname(filePath: string): string {
  const base = getBasename(filePath);
  const idx = base.lastIndexOf('.');
  return idx !== -1 ? base.slice(idx) : '';
}
