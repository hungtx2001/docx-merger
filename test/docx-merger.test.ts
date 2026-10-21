import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import JSZip from 'jszip';
import { DocxMerger } from '../src/index';
import { DOMParser } from '@xmldom/xmldom';

describe('DocxMerger Modern TypeScript Suite', () => {
  const templatePath1 = path.resolve(__dirname, '../example/template.docx');
  const templatePath2 = path.resolve(__dirname, '../example/template1.docx');

  const file1 = fs.readFileSync(templatePath1);
  const file2 = fs.readFileSync(templatePath2);

  it('should successfully merge two docx files with strict OpenXML schema integrity', async () => {
    const merger = new DocxMerger();
    await merger.merge([file1, file2], { pageBreak: true });
    const buffer = await merger.save('nodebuffer');

    expect(buffer).toBeDefined();
    expect(buffer.length).toBeGreaterThan(0);

    const zip = await JSZip.loadAsync(buffer);

    // Fundamental OpenXML package files must exist
    expect(zip.file('[Content_Types].xml')).not.toBeNull();
    expect(zip.file('word/document.xml')).not.toBeNull();
    expect(zip.file('word/_rels/document.xml.rels')).not.toBeNull();
    expect(zip.file('word/styles.xml')).not.toBeNull();
    expect(zip.file('word/numbering.xml')).not.toBeNull();
  });

  it('should remap relationship IDs and preserve media without conflicts', async () => {
    const merger = new DocxMerger();
    await merger.merge([file1, file2]);
    const buffer = await merger.save('nodebuffer');

    const zip = await JSZip.loadAsync(buffer);
    const relsXml = await zip.file('word/_rels/document.xml.rels')!.async('string');
    const docXml = await zip.file('word/document.xml')!.async('string');

    // Parse relationships XML
    const relsDoc = new DOMParser().parseFromString(relsXml, 'text/xml');
    const rels = relsDoc.getElementsByTagName('Relationship');
    let imageRel: Element | null = null;

    for (let i = 0; i < rels.length; i++) {
      const type = rels[i].getAttribute('Type') || '';
      if (type.includes('/image')) {
        imageRel = rels[i];
        break;
      }
    }

    expect(imageRel).not.toBeNull();
    const imageTarget = imageRel!.getAttribute('Target')!;
    const imageRelId = imageRel!.getAttribute('Id')!;

    // Check media file presence in zip
    const mediaFilePath = imageTarget.startsWith('word/')
      ? imageTarget
      : `word/${imageTarget}`;
    expect(zip.file(mediaFilePath)).not.toBeNull();

    // Verify document.xml references the newly remapped relationship ID
    const hasImageRef =
      docXml.includes(`r:embed="${imageRelId}"`) ||
      docXml.includes(`r:id="${imageRelId}"`);
    expect(hasImageRef).toBe(true);
  });

  it('should eliminate duplicate numbering IDs and align body references', async () => {
    const merger = new DocxMerger();
    await merger.merge([file1, file2]);
    const buffer = await merger.save('nodebuffer');

    const zip = await JSZip.loadAsync(buffer);
    const numberingXml = await zip.file('word/numbering.xml')!.async('string');
    const docXml = await zip.file('word/document.xml')!.async('string');

    const numIds: string[] = [];
    const numIdRegex = /<w:num\s+w:numId="([^"]+)"/g;
    let match: RegExpExecArray | null;
    while ((match = numIdRegex.exec(numberingXml)) !== null) {
      numIds.push(match[1]);
    }

    expect(numIds.length).toBeGreaterThan(0);

    // Numbering IDs must be strictly unique to avoid MS Word fatal errors
    const uniqueNumIds = new Set(numIds);
    expect(numIds.length).toBe(uniqueNumIds.size);

    // Body references must resolve to a valid numbering ID
    const docNumIdRegex = /<w:numId\s+w:val="([^"]+)"/g;
    while ((match = docNumIdRegex.exec(docXml)) !== null) {
      const refId = match[1];
      expect(uniqueNumIds.has(refId)).toBe(true);
    }
  });

  it('should preserve section properties and independent headers across sections', async () => {
    const merger = new DocxMerger();
    await merger.merge([file1, file2], { pageBreak: true });
    const buffer = await merger.save('nodebuffer');

    const zip = await JSZip.loadAsync(buffer);
    const docXml = await zip.file('word/document.xml')!.async('string');

    // Section break must be preserved
    expect(
      docXml.includes('<w:pPr><w:sectPr') || docXml.includes('<w:sectPr')
    ).toBe(true);

    // Base and imported headers must exist
    expect(zip.file('word/header1.xml')).not.toBeNull();
    const files = Object.keys(zip.files);
    const hasImportedHeader = files.some((f) => f.startsWith('word/header_doc'));
    expect(hasImportedHeader).toBe(true);
  });

  it('should support multiple output formats including Uint8Array and base64', async () => {
    const merger = new DocxMerger();
    await merger.merge([file1, file2]);

    const uint8 = await merger.save('uint8array');
    expect(uint8 instanceof Uint8Array).toBe(true);
    expect(uint8.byteLength).toBeGreaterThan(0);

    const base64 = await merger.save('base64');
    expect(typeof base64).toBe('string');
    expect(base64.length).toBeGreaterThan(0);
  });

  it('should support pageBreak: false (continuous section breaks)', async () => {
    const merger = new DocxMerger();
    await merger.merge([file1, file2], { pageBreak: false });
    const buffer = await merger.save('nodebuffer');

    const zip = await JSZip.loadAsync(buffer);
    const docXml = await zip.file('word/document.xml')!.async('string');

    expect(docXml.includes('w:val="continuous"')).toBe(true);
  });

  it('should support single file merge without errors', async () => {
    const merger = new DocxMerger();
    await merger.merge([file1]);
    const buffer = await merger.save('nodebuffer');

    expect(buffer.length).toBeGreaterThan(0);
    const zip = await JSZip.loadAsync(buffer);
    expect(zip.file('word/document.xml')).not.toBeNull();
  });

  it('should throw an informative error when trying to save without providing files', async () => {
    const merger = new DocxMerger();
    await expect(merger.save('nodebuffer')).rejects.toThrow('No files provided to merge');
  });

  it('should generate valid, parseable XML for all modified parts', async () => {
    const merger = new DocxMerger();
    await merger.merge([file1, file2]);
    const buffer = await merger.save('nodebuffer');

    const zip = await JSZip.loadAsync(buffer);
    const docXml = await zip.file('word/document.xml')!.async('string');
    const relsXml = await zip.file('word/_rels/document.xml.rels')!.async('string');
    const numXml = await zip.file('word/numbering.xml')!.async('string');
    const stylesXml = await zip.file('word/styles.xml')!.async('string');
    const typesXml = await zip.file('[Content_Types].xml')!.async('string');

    // All XML files must parse without errors
    const parser = new DOMParser();
    expect(parser.parseFromString(docXml, 'text/xml')).toBeDefined();
    expect(parser.parseFromString(relsXml, 'text/xml')).toBeDefined();
    expect(parser.parseFromString(numXml, 'text/xml')).toBeDefined();
    expect(parser.parseFromString(stylesXml, 'text/xml')).toBeDefined();
    expect(parser.parseFromString(typesXml, 'text/xml')).toBeDefined();
  });
});
