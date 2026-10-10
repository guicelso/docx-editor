/**
 * A block-level content control for the browser acceptance of block tags: two paragraphs wrapped
 * by a control between two plain ones, each control tagged with the name its chips show.
 *
 * Run: bun e2e/fixtures/generate-block-content-control-tags-fixture.ts
 */

import JSZip from 'jszip';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { fileURLToPath } from 'node:url';

const FIXTURES_DIR = path.dirname(fileURLToPath(import.meta.url));
const OUT = path.join(FIXTURES_DIR, 'block-content-control-tags.docx');
const ZIP_DATE = new Date('2026-10-10T12:00:00Z');

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const paragraph = (text: string, size = '') =>
  `<w:p><w:r>${size ? `<w:rPr><w:sz w:val="${size}"/></w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const block = (tag: string, inner: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w:richText/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}">
  <w:body>
    ${paragraph('Before.')}
    ${block('block-1', paragraph('first line', '36') + paragraph('second line', '36'))}
    ${paragraph('After.')}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1296" w:right="1296" w:bottom="1296" w:left="1296" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;

export async function createBlockContentControlTagsFixture(): Promise<Uint8Array> {
  const zip = new JSZip();
  const opts = { date: ZIP_DATE, createFolders: false };
  zip.file('[Content_Types].xml', contentTypesXml, opts);
  zip.file('_rels/.rels', relsXml, opts);
  zip.file('word/document.xml', documentXml, opts);
  return zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 9 },
  });
}

if (import.meta.main) {
  const bytes = await createBlockContentControlTagsFixture();
  fs.writeFileSync(OUT, bytes);
  console.log(`Created ${OUT}`);
}
