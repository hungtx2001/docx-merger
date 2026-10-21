import JSZip from 'jszip';
import type {
  DocxMergerOptions,
  InputFileType,
  SaveOutputFormat,
} from '../types';
import { parseXml, serializeXml, getDirectChildren } from '../utils/xml';
import { ContentTypesManager } from './content-types';
import { NumberingManager } from './numbering';
import { StylesManager } from './styles';
import { RelationsManager } from './relations';

export class DocxMerger {
  private pageBreak: boolean;
  private baseStyle: 'source' | 'target';
  private zipList: JSZip[] = [];
  private targetZip: JSZip | null = null;
  private isMerged = false;
  private pendingFiles?: InputFileType[];

  constructor(options: DocxMergerOptions = {}, files: InputFileType[] = []) {
    this.pageBreak = options.pageBreak !== undefined ? !!options.pageBreak : true;
    this.baseStyle = options.style || 'source';

    if (files && files.length > 0) {
      this.pendingFiles = files;
    }
  }

  /**
   * Initializes and executes document merge. (Backwards-compatible API).
   */
  async initialize(
    options: DocxMergerOptions = {},
    files: InputFileType[] = []
  ): Promise<void> {
    this.pageBreak = options.pageBreak !== undefined ? !!options.pageBreak : this.pageBreak;
    this.baseStyle = options.style || this.baseStyle;

    const inputFiles = files.length > 0 ? files : this.pendingFiles || [];
    this.zipList = [];

    for (const file of inputFiles) {
      const zip = await new JSZip().loadAsync(file as any);
      this.zipList.push(zip);
    }

    if (this.zipList.length > 0) {
      await this.executeMerge();
    }
  }

  /**
   * Fluent merge API for merging documents with options.
   */
  async merge(
    files: InputFileType[],
    options: DocxMergerOptions = {}
  ): Promise<DocxMerger> {
    await this.initialize(options, files);
    return this;
  }

  private async executeMerge(): Promise<void> {
    if (this.zipList.length === 0) return;

    this.targetZip = this.zipList[0];

    const contentTypesManager = new ContentTypesManager();
    await contentTypesManager.load(this.targetZip);

    const numberingManager = new NumberingManager();
    await numberingManager.load(this.targetZip);

    const stylesManager = new StylesManager();
    await stylesManager.load(this.targetZip);

    const relationsManager = new RelationsManager();
    await relationsManager.load(this.targetZip);

    // Load base document.xml
    const baseDocFile = this.targetZip.file('word/document.xml');
    if (!baseDocFile) {
      throw new Error('Invalid DOCX: missing word/document.xml in base file.');
    }

    const baseDocXmlString = await baseDocFile.async('string');
    const baseDoc = parseXml(baseDocXmlString);
    if (!baseDoc) {
      throw new Error('Failed to parse base document.xml');
    }

    const baseBody = baseDoc.getElementsByTagName('w:body')[0];
    if (!baseBody) {
      throw new Error('Invalid DOCX: missing <w:body> in base document.');
    }

    const mergedBodyElements: Element[] = [];
    let finalSectPr: Element | null = null;

    for (let i = 0; i < this.zipList.length; i++) {
      const currentZip = this.zipList[i];
      const isBase = i === 0;
      const isLast = i === this.zipList.length - 1;

      let docXmlDom: Document | null;
      if (isBase) {
        docXmlDom = baseDoc;
      } else {
        await contentTypesManager.merge(currentZip);

        const currentDocFile = currentZip.file('word/document.xml');
        if (!currentDocFile) continue;

        const xmlString = await currentDocFile.async('string');
        docXmlDom = parseXml(xmlString);
        if (!docXmlDom) continue;

        await numberingManager.mergeFileNumbering(currentZip, docXmlDom, i);
        await stylesManager.mergeFileStyles(currentZip, docXmlDom, i, this.baseStyle);
        await relationsManager.mergeFileRelations(
          currentZip,
          this.targetZip,
          docXmlDom,
          i,
          contentTypesManager
        );
      }

      const currentBody = docXmlDom.getElementsByTagName('w:body')[0];
      if (!currentBody) continue;

      const children = getDirectChildren(currentBody);
      const contentNodes: Element[] = [];
      let currentSectPr: Element | null = null;

      for (const child of children) {
        const localName = child.localName || child.nodeName;
        if (localName === 'sectPr' || localName.endsWith(':sectPr')) {
          currentSectPr = child;
        } else {
          contentNodes.push(child);
        }
      }

      if (!isLast) {
        // Detect if this document already has an inner section break at or near the end
        let lastSectPrIndex = -1;
        for (let k = 0; k < contentNodes.length; k++) {
          if (contentNodes[k].getElementsByTagName('w:sectPr').length > 0) {
            lastSectPrIndex = k;
          }
        }

        if (lastSectPrIndex !== -1) {
          // Check if all nodes after lastSectPrIndex are empty paragraphs
          let allEmptyAfter = true;
          for (let k = lastSectPrIndex + 1; k < contentNodes.length; k++) {
            if (!this.isParagraphEmpty(contentNodes[k])) {
              allEmptyAfter = false;
              break;
            }
          }

          if (allEmptyAfter) {
            // Truncate redundant empty paragraphs after the existing section break
            contentNodes.splice(lastSectPrIndex + 1);

            // Ensure the existing section break has the appropriate w:type
            const sectInP = contentNodes[lastSectPrIndex].getElementsByTagName('w:sectPr')[0];
            if (sectInP && docXmlDom) {
              this.configureSectionBreak(sectInP, this.pageBreak, docXmlDom);
            }
            currentSectPr = null; // Do not append a second section break (which causes an empty blank page)
          }
        }

        // Trim trailing empty paragraphs if currentSectPr is still to be attached
        if (currentSectPr) {
          while (
            contentNodes.length > 0 &&
            this.isParagraphEmpty(contentNodes[contentNodes.length - 1])
          ) {
            contentNodes.pop();
          }
        }
      }

      // Append content nodes
      for (const node of contentNodes) {
        const importedNode = isBase
          ? node
          : (baseDoc.importNode
              ? (baseDoc.importNode(node, true) as Element)
              : (node.cloneNode(true) as Element));
        mergedBodyElements.push(importedNode);
      }

      if (!isLast && currentSectPr) {
        // Attach currentSectPr directly to the last paragraph if possible, avoiding empty paragraphs
        const lastElement =
          mergedBodyElements.length > 0
            ? mergedBodyElements[mergedBodyElements.length - 1]
            : null;
        const clonedSectPr = isBase
          ? (currentSectPr.cloneNode(true) as Element)
          : (baseDoc.importNode
              ? (baseDoc.importNode(currentSectPr, true) as Element)
              : (currentSectPr.cloneNode(true) as Element));

        this.configureSectionBreak(clonedSectPr, this.pageBreak, baseDoc);

        if (
          lastElement &&
          (lastElement.localName === 'p' || lastElement.nodeName === 'w:p')
        ) {
          let pPr = lastElement.getElementsByTagName('w:pPr')[0];
          if (!pPr) {
            pPr = baseDoc.createElement('w:pPr');
            lastElement.insertBefore(pPr, lastElement.firstChild);
          }
          pPr.appendChild(clonedSectPr);
        } else {
          const p = baseDoc.createElement('w:p');
          const pPr = baseDoc.createElement('w:pPr');
          pPr.appendChild(clonedSectPr);
          p.appendChild(pPr);
          mergedBodyElements.push(p);
        }
      } else if (isLast) {
        // Last file's sectPr becomes the root sectPr directly under <w:body>
        if (currentSectPr) {
          finalSectPr = isBase
            ? currentSectPr
            : (baseDoc.importNode
                ? (baseDoc.importNode(currentSectPr, true) as Element)
                : (currentSectPr.cloneNode(true) as Element));
        }
      }
    }

    // Reconstruct <w:body>
    while (baseBody.firstChild) {
      baseBody.removeChild(baseBody.firstChild);
    }

    for (const el of mergedBodyElements) {
      baseBody.appendChild(el);
    }

    if (finalSectPr) {
      baseBody.appendChild(finalSectPr);
    }

    // Save parts to targetZip
    this.targetZip.file('word/document.xml', serializeXml(baseDoc));
    contentTypesManager.save(this.targetZip);
    numberingManager.save(this.targetZip, contentTypesManager, relationsManager);
    stylesManager.save(this.targetZip);
    relationsManager.save(this.targetZip);

    this.isMerged = true;
  }

  private isParagraphEmpty(p: Element): boolean {
    if (
      p.getElementsByTagName('w:drawing').length > 0 ||
      p.getElementsByTagName('w:pict').length > 0 ||
      p.getElementsByTagName('w:object').length > 0
    ) {
      return false;
    }
    return (p.textContent || '').trim().length === 0;
  }

  private configureSectionBreak(
    sectPr: Element,
    pageBreak: boolean,
    doc: Document
  ): void {
    let typeNode = sectPr.getElementsByTagName('w:type')[0];
    const val = pageBreak ? 'nextPage' : 'continuous';
    if (!typeNode) {
      typeNode = doc.createElement('w:type');
      // In ECMA-376 schema, <w:type> is positioned before <w:pgSz> and after header/footer references
      const pgSz = sectPr.getElementsByTagName('w:pgSz')[0];
      if (pgSz) {
        sectPr.insertBefore(typeNode, pgSz);
      } else {
        sectPr.appendChild(typeNode);
      }
    }
    typeNode.setAttribute('w:val', val);
  }

  /**
   * Generates and returns the merged document in the requested format.
   * @param type Output format (default: 'nodebuffer')
   */
  async save(type?: 'nodebuffer'): Promise<Buffer>;
  async save(type: 'blob'): Promise<Blob>;
  async save(type: 'arraybuffer'): Promise<ArrayBuffer>;
  async save(type: 'uint8array'): Promise<Uint8Array>;
  async save(type: 'base64' | 'binarystring'): Promise<string>;
  async save(type?: SaveOutputFormat): Promise<any>;
  async save(type: SaveOutputFormat = 'nodebuffer'): Promise<any> {
    if (!this.isMerged) {
      if (this.pendingFiles && this.pendingFiles.length > 0) {
        await this.initialize({}, this.pendingFiles);
      } else if (this.zipList.length > 0) {
        await this.executeMerge();
      } else {
        throw new Error('No files provided to merge.');
      }
    }

    if (!this.targetZip) {
      throw new Error('No target document available.');
    }

    return await this.targetZip.generateAsync({
      type: type as any,
      compression: 'DEFLATE',
      compressionOptions: {
        level: 4,
      },
    });
  }
}

export default DocxMerger;
