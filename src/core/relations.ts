import type JSZip from 'jszip';
import { parseXml, serializeXml, walkElements, getBasename, getExtname } from '../utils/xml';
import { RELATIONSHIP_TYPES, CONTENT_TYPES, MIME_MAP } from '../utils/constants';
import type { ContentTypesManager } from './content-types';

export class RelationsManager {
  private doc: Document | null = null;
  private maxRIdIndex = 0;
  private existingRels = new Map<string, Element>();

  async load(baseZip: JSZip): Promise<void> {
    const file = baseZip.file('word/_rels/document.xml.rels');
    if (!file) {
      this.ensureBaseDoc();
      return;
    }

    const xmlString = await file.async('string');
    this.doc = parseXml(xmlString);
    if (!this.doc) {
      this.ensureBaseDoc();
      return;
    }

    const relNodes = this.doc.getElementsByTagName('Relationship');
    for (let i = 0; i < relNodes.length; i++) {
      const id = relNodes[i].getAttribute('Id');
      if (id) {
        this.existingRels.set(id, relNodes[i]);
        const match = id.match(/rId(\d+)/i);
        if (match) {
          const num = parseInt(match[1], 10);
          if (num > this.maxRIdIndex) {
            this.maxRIdIndex = num;
          }
        }
      }
    }
  }

  private ensureBaseDoc(): void {
    if (!this.doc) {
      const template = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`;
      this.doc = parseXml(template);
    }
  }

  generateNextRId(): string {
    this.maxRIdIndex++;
    return `rId${this.maxRIdIndex}`;
  }

  ensureMainRelationship(type: string, target: string): string {
    this.ensureBaseDoc();
    if (!this.doc) return '';

    const root = this.doc.documentElement;
    const rels = root.getElementsByTagName('Relationship');
    for (let i = 0; i < rels.length; i++) {
      if (
        rels[i].getAttribute('Type') === type &&
        rels[i].getAttribute('Target') === target
      ) {
        return rels[i].getAttribute('Id') || '';
      }
    }

    const newId = this.generateNextRId();
    const rel = this.doc.createElement('Relationship');
    rel.setAttribute('Id', newId);
    rel.setAttribute('Type', type);
    rel.setAttribute('Target', target);
    root.appendChild(rel);
    this.existingRels.set(newId, rel);
    return newId;
  }

  async mergeFileRelations(
    sourceZip: JSZip,
    targetZip: JSZip,
    sourceDocXmlDom: Document | null,
    fileIndex: number,
    contentTypesManager?: ContentTypesManager
  ): Promise<void> {
    if (fileIndex === 0) return; // Base relationships already loaded

    const relFile = sourceZip.file('word/_rels/document.xml.rels');
    if (!relFile) return;

    const xmlString = await relFile.async('string');
    const sourceRelsDoc = parseXml(xmlString);
    if (!sourceRelsDoc) return;

    this.ensureBaseDoc();
    if (!this.doc) return;

    const baseRoot = this.doc.documentElement;
    const relNodes = sourceRelsDoc.getElementsByTagName('Relationship');
    const relMap = new Map<string, string>(); // oldRId -> newRId

    for (let i = 0; i < relNodes.length; i++) {
      const node = relNodes[i];
      const oldId = node.getAttribute('Id');
      const type = node.getAttribute('Type');
      const target = node.getAttribute('Target') || '';
      const targetMode = node.getAttribute('TargetMode');

      if (!oldId || !type) continue;

      // Skip document-level shared parts as they are provided by the primary base document
      if (
        type === RELATIONSHIP_TYPES.styles ||
        type === RELATIONSHIP_TYPES.numbering ||
        type === RELATIONSHIP_TYPES.fontTable ||
        type === RELATIONSHIP_TYPES.theme ||
        type === RELATIONSHIP_TYPES.settings ||
        type.includes('/webSettings') ||
        type.includes('/stylesWithEffects')
      ) {
        continue;
      }

      // If internal part is not referenced in the document body and is not header/footer, skip orphan relationship
      const isHeaderOrFooter =
        type === RELATIONSHIP_TYPES.header ||
        type === RELATIONSHIP_TYPES.footer ||
        target.includes('header') ||
        target.includes('footer');

      if (!isHeaderOrFooter && targetMode !== 'External' && !this.isRelReferenced(sourceDocXmlDom, oldId)) {
        continue;
      }

      const newId = this.generateNextRId();
      relMap.set(oldId, newId);

      let newTarget = target;

      // 1. Media (Images)
      if (type === RELATIONSHIP_TYPES.image || target.includes('media/')) {
        const cleanTarget = target.replace(/^word\//, '').replace(/^\.\.\//, '');
        const filename = getBasename(cleanTarget);
        const ext = getExtname(filename).toLowerCase().replace('.', '');
        const newFilename = `doc${fileIndex}_${this.maxRIdIndex}_${filename}`;
        newTarget = `media/${newFilename}`;

        const mediaSourcePath = target.startsWith('word/')
          ? target
          : `word/${cleanTarget}`;
        const mediaFile =
          sourceZip.file(mediaSourcePath) || sourceZip.file(cleanTarget);
        if (mediaFile) {
          const data = await mediaFile.async('uint8array');
          targetZip.file(`word/${newTarget}`, data);
        }

        if (contentTypesManager && ext && MIME_MAP[ext]) {
          contentTypesManager.ensureDefault(ext, MIME_MAP[ext]);
        }
      } else if (type === RELATIONSHIP_TYPES.header || target.includes('header')) {
        // 2. Headers
        const cleanTarget = target.replace(/^word\//, '');
        const newFilename = `header_doc${fileIndex}_${this.maxRIdIndex}.xml`;
        newTarget = newFilename;

        const headerSourcePath = `word/${cleanTarget}`;
        const headerFile = sourceZip.file(headerSourcePath);
        if (headerFile) {
          const headerXml = await headerFile.async('string');
          const headerRelsPath = `word/_rels/${cleanTarget}.rels`;
          const headerRelsFile = sourceZip.file(headerRelsPath);
          if (headerRelsFile) {
            const headerRelsXml = await headerRelsFile.async('string');
            targetZip.file(`word/_rels/${newFilename}.rels`, headerRelsXml);
          }
          targetZip.file(`word/${newFilename}`, headerXml);

          if (contentTypesManager) {
            contentTypesManager.ensureOverride(
              `/word/${newFilename}`,
              CONTENT_TYPES.header
            );
          }
        }
      } else if (type === RELATIONSHIP_TYPES.footer || target.includes('footer')) {
        // 3. Footers
        const cleanTarget = target.replace(/^word\//, '');
        const newFilename = `footer_doc${fileIndex}_${this.maxRIdIndex}.xml`;
        newTarget = newFilename;

        const footerSourcePath = `word/${cleanTarget}`;
        const footerFile = sourceZip.file(footerSourcePath);
        if (footerFile) {
          const footerXml = await footerFile.async('string');
          const footerRelsPath = `word/_rels/${cleanTarget}.rels`;
          const footerRelsFile = sourceZip.file(footerRelsPath);
          if (footerRelsFile) {
            const footerRelsXml = await footerRelsFile.async('string');
            targetZip.file(`word/_rels/${newFilename}.rels`, footerRelsXml);
          }
          targetZip.file(`word/${newFilename}`, footerXml);

          if (contentTypesManager) {
            contentTypesManager.ensureOverride(
              `/word/${newFilename}`,
              CONTENT_TYPES.footer
            );
          }
        }
      } else if (target && targetMode !== 'External') {
        // 4. Other internal parts (customXml, comments, footnotes, etc.)
        const cleanTarget = target.replace(/^word\//, '').replace(/^\.\.\//, '');
        const sourcePath = target.startsWith('word/')
          ? target
          : target.startsWith('../')
            ? cleanTarget
            : `word/${cleanTarget}`;
        const partFile = sourceZip.file(sourcePath);
        if (partFile) {
          const data = await partFile.async('uint8array');
          targetZip.file(sourcePath, data);
        }
      }

      // Append new Relationship node
      const newRelNode = this.doc.createElement('Relationship');
      newRelNode.setAttribute('Id', newId);
      newRelNode.setAttribute('Type', type);
      newRelNode.setAttribute('Target', newTarget);
      if (targetMode) {
        newRelNode.setAttribute('TargetMode', targetMode);
      }
      baseRoot.appendChild(newRelNode);
      this.existingRels.set(newId, newRelNode);
    }

    // Update all relationship references in source document.xml DOM
    if (sourceDocXmlDom && relMap.size > 0) {
      walkElements(sourceDocXmlDom, (el) => {
        if (!el.attributes) return;
        for (let a = 0; a < el.attributes.length; a++) {
          const attr = el.attributes[a];
          const name = attr.name;
          if (
            name === 'r:id' ||
            name === 'r:embed' ||
            name === 'r:link' ||
            name === 'r:href' ||
            name.endsWith(':id') ||
            name.endsWith(':embed') ||
            name.endsWith(':link') ||
            name.endsWith(':href')
          ) {
            const val = attr.value;
            if (relMap.has(val)) {
              const newVal = relMap.get(val)!;
              attr.value = newVal;
              el.setAttribute(name, newVal);
            }
          }
        }
      });
    }
  }

  private isRelReferenced(doc: Document | null, relId: string): boolean {
    if (!doc) return false;
    let found = false;
    walkElements(doc, (el) => {
      if (found || !el.attributes) return;
      for (let a = 0; a < el.attributes.length; a++) {
        if (el.attributes[a].value === relId) {
          found = true;
          return;
        }
      }
    });
    return found;
  }

  save(targetZip: JSZip): void {
    if (!this.doc) return;
    targetZip.file('word/_rels/document.xml.rels', serializeXml(this.doc));
  }
}
