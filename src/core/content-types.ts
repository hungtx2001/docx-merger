import type JSZip from 'jszip';
import { parseXml, serializeXml } from '../utils/xml';

interface OverrideEntry {
  originalPartName: string;
  contentType: string;
}

export class ContentTypesManager {
  private defaults = new Map<string, string>(); // ext -> contentType
  private overrides = new Map<string, OverrideEntry>(); // normalized partName -> entry
  private doc: Document | null = null;

  async load(baseZip: JSZip): Promise<void> {
    const file = baseZip.file('[Content_Types].xml');
    if (!file) return;

    const xmlString = await file.async('string');
    this.doc = parseXml(xmlString);
    if (!this.doc) return;

    const defaultNodes = this.doc.getElementsByTagName('Default');
    for (let i = 0; i < defaultNodes.length; i++) {
      const ext = defaultNodes[i].getAttribute('Extension');
      const ct = defaultNodes[i].getAttribute('ContentType');
      if (ext && ct) {
        this.defaults.set(ext.toLowerCase(), ct);
      }
    }

    const overrideNodes = this.doc.getElementsByTagName('Override');
    for (let i = 0; i < overrideNodes.length; i++) {
      const part = overrideNodes[i].getAttribute('PartName');
      const ct = overrideNodes[i].getAttribute('ContentType');
      if (part && ct) {
        this.overrides.set(this.normalizePartName(part), {
          originalPartName: part,
          contentType: ct,
        });
      }
    }
  }

  private normalizePartName(partName: string): string {
    let p = partName.trim();
    if (!p.startsWith('/')) p = '/' + p;
    return p.toLowerCase();
  }

  ensureDefault(extension: string, contentType: string): void {
    if (!extension || !contentType) return;
    const ext = extension.toLowerCase().replace(/^\./, '');
    if (!this.defaults.has(ext)) {
      this.defaults.set(ext, contentType);
    }
  }

  ensureOverride(partName: string, contentType: string): void {
    if (!partName || !contentType) return;
    let formatted = partName.trim();
    if (!formatted.startsWith('/')) formatted = '/' + formatted;
    this.overrides.set(this.normalizePartName(formatted), {
      originalPartName: formatted,
      contentType,
    });
  }

  async merge(sourceZip: JSZip): Promise<void> {
    const file = sourceZip.file('[Content_Types].xml');
    if (!file) return;

    const xmlString = await file.async('string');
    const sourceDoc = parseXml(xmlString);
    if (!sourceDoc) return;

    const defaultNodes = sourceDoc.getElementsByTagName('Default');
    for (let i = 0; i < defaultNodes.length; i++) {
      const ext = defaultNodes[i].getAttribute('Extension');
      const ct = defaultNodes[i].getAttribute('ContentType');
      if (ext && ct) {
        this.ensureDefault(ext, ct);
      }
    }

    const overrideNodes = sourceDoc.getElementsByTagName('Override');
    for (let i = 0; i < overrideNodes.length; i++) {
      const part = overrideNodes[i].getAttribute('PartName');
      const ct = overrideNodes[i].getAttribute('ContentType');
      if (part && ct) {
        const normalized = this.normalizePartName(part);
        if (!this.overrides.has(normalized)) {
          this.overrides.set(normalized, {
            originalPartName: part,
            contentType: ct,
          });
        }
      }
    }
  }

  save(targetZip: JSZip): void {
    if (!this.doc) {
      this.doc = parseXml(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>'
      );
    }
    if (!this.doc) return;

    const root = this.doc.documentElement;
    while (root.firstChild) {
      root.removeChild(root.firstChild);
    }

    // Always ensure essential default extensions exist
    this.ensureDefault('rels', 'application/vnd.openxmlformats-package.relationships+xml');
    this.ensureDefault('xml', 'application/xml');

    for (const [ext, ct] of this.defaults.entries()) {
      const el = this.doc.createElement('Default');
      el.setAttribute('Extension', ext);
      el.setAttribute('ContentType', ct);
      root.appendChild(el);
    }

    for (const { originalPartName, contentType } of this.overrides.values()) {
      const zipPath = originalPartName.replace(/^\//, '');

      // Rule 1: Per ECMA-376 Part 2, .rels parts must NEVER have an <Override> element
      if (zipPath.endsWith('.rels')) {
        continue;
      }

      // Rule 2: Media files should be handled by <Default Extension="..."> rather than <Override>
      const dotIdx = zipPath.lastIndexOf('.');
      const ext = dotIdx !== -1 ? zipPath.slice(dotIdx + 1).toLowerCase() : '';
      if (['jpeg', 'jpg', 'png', 'gif', 'bmp', 'svg', 'emf', 'wmf'].includes(ext)) {
        this.ensureDefault(ext, contentType);
        continue;
      }

      // Rule 3: Only emit <Override> for parts that actually exist in the target archive
      if (!targetZip.file(zipPath)) {
        continue;
      }

      const el = this.doc.createElement('Override');
      el.setAttribute('PartName', originalPartName);
      el.setAttribute('ContentType', contentType);
      root.appendChild(el);
    }

    targetZip.file('[Content_Types].xml', serializeXml(this.doc));
  }
}
