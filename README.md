# Docx-merger

A fast, reliable JavaScript/TypeScript library for merging DOCX files with full OpenXML schema compliance, preserving styles, numbering/bullets, images/media, headers/footers, and section orientations.

Works seamlessly in **Node.js**, **modern browsers**, and edge environments.

---

## ✨ Features

- **Reliable OpenXML Compliance:** Remaps all Relationship IDs (`rId`), Numbering IDs (`numId`, `abstractNumId`), and media assets without collisions or XML schema corruption.
- **Full Media & Image Preservation:** Retains embedded images, figures, and charts across all merged documents.
- **Preserves Numbering & Bullets:** Eliminates duplicate `w:numId` tags that trigger Microsoft Word's *"unreadable content"* error.
- **Section & Layout Preservation:** Preserves page orientations (Portrait / Landscape), margins, and headers/footers per document.
- **Browser & Node.js Compatible:** Zero Node.js-only core dependencies (uses browser-safe path parsing and pure JSZip/XML-DOM).
- **TypeScript Support:** Includes complete TypeScript definitions out of the box (`.d.ts`).
- **Modern Async/Await API:** Fast, promise-based API with backwards-compatible methods.

---

## 📦 Installation

```bash
npm install @hungtrinh2001/docx-merger
```

---

## 🚀 Usage

### ESM / TypeScript (Recommended)

```typescript
import DocxMerger from '@hungtrinh2001/docx-merger';
// or: import { DocxMerger } from '@hungtrinh2001/docx-merger';
import * as fs from 'fs';

async function mergeDocs() {
  const file1 = fs.readFileSync('document1.docx');
  const file2 = fs.readFileSync('document2.docx');

  const merger = new DocxMerger();
  
  // Merge multiple files with options
  await merger.merge([file1, file2], {
    pageBreak: true // Insert page break between files (default: true)
  });

  // Save as Node.js Buffer
  const outputBuffer = await merger.save('nodebuffer');
  fs.writeFileSync('merged.docx', outputBuffer);
}
```

### CommonJS (CJS)

```javascript
const { DocxMerger } = require('@hungtrinh2001/docx-merger');
const fs = require('fs');

(async () => {
  const file1 = fs.readFileSync('document1.docx');
  const file2 = fs.readFileSync('document2.docx');

  const merger = new DocxMerger();
  await merger.merge([file1, file2], { pageBreak: true });

  const outputBuffer = await merger.save('nodebuffer');
  fs.writeFileSync('merged.docx', outputBuffer);
})();
```

### Browser Usage

```javascript
import DocxMerger from '@hungtrinh2001/docx-merger';

async function handleMerge(fileBlob1, fileBlob2) {
  const buffer1 = await fileBlob1.arrayBuffer();
  const buffer2 = await fileBlob2.arrayBuffer();

  const merger = new DocxMerger();
  await merger.merge([buffer1, buffer2], { pageBreak: true });

  const blob = await merger.save('blob');
  
  // Download or preview the blob in browser
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'merged.docx';
  a.click();
}
```

---

## ⚙️ Options

| Option | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| `pageBreak` | `boolean` | `true` | Inserts a section page break between merged documents. If `false`, documents flow continuously. |
| `style` | `'source' \| 'target'` | `'source'` | Precedence strategy for styles. |

### Supported Output Formats (`save(type)`)

- `'nodebuffer'` (Node.js `Buffer`)
- `'blob'` (Browser `Blob`)
- `'uint8array'` (`Uint8Array`)
- `'arraybuffer'` (`ArrayBuffer`)
- `'base64'` (`string`)

---

## 🧪 Testing

Run the automated unit test suite:

```bash
npm test
```

Build production bundles:

```bash
npm run build
```

---

## 📄 License

Apache-2.0
