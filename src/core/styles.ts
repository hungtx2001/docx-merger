import type JSZip from 'jszip';
import { parseXml, serializeXml, walkElements } from '../utils/xml';
import { BUILT_IN_STYLES } from '../utils/constants';

export class StylesManager {
  private doc: Document | null = null;
  private existingStyleIds = new Set<string>();

  async load(baseZip: JSZip): Promise<void> {
    const file = baseZip.file('word/styles.xml');
    if (!file) return;

    const xmlString = await file.async('string');
    this.doc = parseXml(xmlString);
    if (!this.doc) return;

    const styleNodes = this.doc.getElementsByTagName('w:style');
    for (let i = 0; i < styleNodes.length; i++) {
      const id = styleNodes[i].getAttribute('w:styleId');
      if (id) {
        this.existingStyleIds.add(id);
      }
    }
  }

  async mergeFileStyles(
    sourceZip: JSZip,
    sourceDocXmlDom: Document | null,
    fileIndex: number,
    strategy: 'source' | 'target' = 'source'
  ): Promise<void> {
    if (fileIndex === 0) return; // Base styles already loaded
    const file = sourceZip.file('word/styles.xml');
    if (!file || !this.doc) return;

    const xmlString = await file.async('string');
    const sourceStylesDoc = parseXml(xmlString);
    if (!sourceStylesDoc) return;

    const baseRoot = this.doc.documentElement;
    const styleNodes = sourceStylesDoc.getElementsByTagName('w:style');
    const styleMap = new Map<string, string>(); // oldStyleId -> newStyleId

    for (let i = 0; i < styleNodes.length; i++) {
      const node = styleNodes[i];
      const styleId = node.getAttribute('w:styleId');
      if (!styleId) continue;

      if (this.existingStyleIds.has(styleId)) {
        // Built-in styles (e.g. Normal, Heading 1) are preserved from base
        if (BUILT_IN_STYLES.has(styleId) || strategy === 'target') {
          continue;
        }

        // Custom style collision -> rename uniquely for 'source' strategy
        const newStyleId = `${styleId}_${fileIndex}`;
        styleMap.set(styleId, newStyleId);
        this.existingStyleIds.add(newStyleId);

        const cloned = (this.doc.importNode
          ? this.doc.importNode(node, true)
          : node.cloneNode(true)) as Element;
        cloned.setAttribute('w:styleId', newStyleId);

        this.updateStyleInternalReferences(cloned, styleMap);
        baseRoot.appendChild(cloned);
      } else {
        // New unique style -> import directly
        this.existingStyleIds.add(styleId);
        const cloned = (this.doc.importNode
          ? this.doc.importNode(node, true)
          : node.cloneNode(true)) as Element;
        baseRoot.appendChild(cloned);
      }
    }

    // Update style references in source document.xml DOM
    if (sourceDocXmlDom && styleMap.size > 0) {
      walkElements(sourceDocXmlDom, (el) => {
        const name = el.localName || el.nodeName;
        if (
          name === 'pStyle' ||
          name === 'rStyle' ||
          name === 'tblStyle' ||
          name.endsWith(':pStyle') ||
          name.endsWith(':rStyle') ||
          name.endsWith(':tblStyle')
        ) {
          const currentVal = el.getAttribute('w:val');
          if (currentVal && styleMap.has(currentVal)) {
            el.setAttribute('w:val', styleMap.get(currentVal)!);
          }
        }
      });
    }
  }

  private updateStyleInternalReferences(
    styleNode: Element,
    styleMap: Map<string, string>
  ): void {
    const updateVal = (tagName: string) => {
      const el = styleNode.getElementsByTagName(tagName)[0];
      if (el) {
        const val = el.getAttribute('w:val');
        if (val && styleMap.has(val)) {
          el.setAttribute('w:val', styleMap.get(val)!);
        }
      }
    };

    updateVal('w:basedOn');
    updateVal('w:next');
    updateVal('w:link');
  }

  save(targetZip: JSZip): void {
    if (!this.doc) return;
    targetZip.file('word/styles.xml', serializeXml(this.doc));
  }
}
