import type JSZip from 'jszip';
import { parseXml, serializeXml, walkElements } from '../utils/xml';
import { NAMESPACES, RELATIONSHIP_TYPES, CONTENT_TYPES } from '../utils/constants';
import type { ContentTypesManager } from './content-types';
import type { RelationsManager } from './relations';

export class NumberingManager {
  private doc: Document | null = null;
  private maxAbstractNumId = 0;
  private maxNumId = 0;
  public hasNumbering = false;

  async load(baseZip: JSZip): Promise<void> {
    const file = baseZip.file('word/numbering.xml');
    if (file) {
      const xmlString = await file.async('string');
      this.doc = parseXml(xmlString);
      if (this.doc) {
        this.hasNumbering = true;
        this.computeMaxIds(this.doc);
      }
    }
  }

  private computeMaxIds(doc: Document): void {
    const abstractNodes = doc.getElementsByTagName('w:abstractNum');
    for (let i = 0; i < abstractNodes.length; i++) {
      const id = parseInt(abstractNodes[i].getAttribute('w:abstractNumId') || '', 10);
      if (!isNaN(id) && id > this.maxAbstractNumId) {
        this.maxAbstractNumId = id;
      }
    }

    const numNodes = doc.getElementsByTagName('w:num');
    for (let i = 0; i < numNodes.length; i++) {
      const id = parseInt(numNodes[i].getAttribute('w:numId') || '', 10);
      if (!isNaN(id) && id > this.maxNumId) {
        this.maxNumId = id;
      }
    }
  }

  private ensureBaseDoc(): void {
    if (!this.doc) {
      const template = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="${NAMESPACES.w}" xmlns:r="${NAMESPACES.r}">
</w:numbering>`;
      this.doc = parseXml(template);
      this.hasNumbering = true;
    }
  }

  async mergeFileNumbering(
    sourceZip: JSZip,
    sourceDocXmlDom: Document | null,
    fileIndex: number
  ): Promise<void> {
    const file = sourceZip.file('word/numbering.xml');
    if (!file) return;

    const xmlString = await file.async('string');
    const sourceNumberingDoc = parseXml(xmlString);
    if (!sourceNumberingDoc) return;

    this.ensureBaseDoc();
    if (!this.doc) return;

    const baseRoot = this.doc.documentElement;
    if (fileIndex === 0) return; // Base numbering already loaded

    const abstractNumMap = new Map<string, string>();
    const numMap = new Map<string, string>();

    // 1. Remap and insert abstractNum elements before any num elements
    const abstractNodes = sourceNumberingDoc.getElementsByTagName('w:abstractNum');
    for (let i = 0; i < abstractNodes.length; i++) {
      const node = abstractNodes[i];
      const oldId = node.getAttribute('w:abstractNumId');
      if (!oldId) continue;

      this.maxAbstractNumId++;
      const newId = String(this.maxAbstractNumId);
      abstractNumMap.set(oldId, newId);

      const cloned = (this.doc.importNode
        ? this.doc.importNode(node, true)
        : node.cloneNode(true)) as Element;
      cloned.setAttribute('w:abstractNumId', newId);

      const firstNumNode = baseRoot.getElementsByTagName('w:num')[0];
      if (firstNumNode) {
        baseRoot.insertBefore(cloned, firstNumNode);
      } else {
        baseRoot.appendChild(cloned);
      }
    }

    // 2. Remap and append num elements
    const numNodes = sourceNumberingDoc.getElementsByTagName('w:num');
    for (let i = 0; i < numNodes.length; i++) {
      const node = numNodes[i];
      const oldId = node.getAttribute('w:numId');
      if (!oldId) continue;

      this.maxNumId++;
      const newId = String(this.maxNumId);
      numMap.set(oldId, newId);

      const cloned = (this.doc.importNode
        ? this.doc.importNode(node, true)
        : node.cloneNode(true)) as Element;
      cloned.setAttribute('w:numId', newId);

      const absRef = cloned.getElementsByTagName('w:abstractNumId')[0];
      if (absRef) {
        const oldAbsId = absRef.getAttribute('w:val');
        if (oldAbsId && abstractNumMap.has(oldAbsId)) {
          absRef.setAttribute('w:val', abstractNumMap.get(oldAbsId)!);
        }
      }

      baseRoot.appendChild(cloned);
    }

    // 3. Update references in source document.xml DOM
    if (sourceDocXmlDom && numMap.size > 0) {
      walkElements(sourceDocXmlDom, (el) => {
        if (
          el.localName === 'numId' ||
          el.nodeName === 'w:numId' ||
          el.nodeName.endsWith(':numId')
        ) {
          const currentVal = el.getAttribute('w:val');
          if (currentVal && numMap.has(currentVal)) {
            el.setAttribute('w:val', numMap.get(currentVal)!);
          }
        }
      });
    }
  }

  save(
    targetZip: JSZip,
    contentTypesManager?: ContentTypesManager,
    relationsManager?: RelationsManager
  ): void {
    if (!this.hasNumbering || !this.doc) return;

    targetZip.file('word/numbering.xml', serializeXml(this.doc));

    if (contentTypesManager) {
      contentTypesManager.ensureOverride('/word/numbering.xml', CONTENT_TYPES.numbering);
    }

    if (relationsManager) {
      relationsManager.ensureMainRelationship(RELATIONSHIP_TYPES.numbering, 'numbering.xml');
    }
  }
}
