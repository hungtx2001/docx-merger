import DocxMerger from '../dist/index.js';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const file1 = fs.readFileSync(path.resolve(__dirname, 'template.docx'));
const file2 = fs.readFileSync(path.resolve(__dirname, 'template1.docx'));

console.log('Merging DOCX files using modern ESM API...');

const merger = new DocxMerger();
await merger.merge([file1, file2], { pageBreak: true });

const outputBuffer = await merger.save('nodebuffer');
fs.writeFileSync(path.resolve(__dirname, 'output.docx'), outputBuffer);

console.log('Successfully merged documents into example/output.docx');
