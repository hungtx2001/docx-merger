const { DocxMerger } = require('../dist/index.cjs');
const fs = require('fs');
const path = require('path');

(async () => {
  console.log('Merging DOCX files using CommonJS (CJS) API...');

  const file1 = fs.readFileSync(path.resolve(__dirname, 'template.docx'));
  const file2 = fs.readFileSync(path.resolve(__dirname, 'template1.docx'));

  const merger = new DocxMerger();
  await merger.merge([file1, file2], { pageBreak: true });

  const outputBuffer = await merger.save('nodebuffer');
  fs.writeFileSync(path.resolve(__dirname, 'output.docx'), outputBuffer);

  console.log('CJS: Successfully merged documents into example/output.docx');
})();
