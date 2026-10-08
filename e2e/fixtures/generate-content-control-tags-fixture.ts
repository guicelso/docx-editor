/**
 * Nested inline content controls for the content-control tag browser acceptance: two groups of
 * two controls in one paragraph, each control tagged with the name its chips show.
 *
 * Run: bun e2e/fixtures/generate-content-control-tags-fixture.ts
 */

import JSZip from 'jszip';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { fileURLToPath } from 'node:url';

const FIXTURES_DIR = path.dirname(fileURLToPath(import.meta.url));
const OUT = path.join(FIXTURES_DIR, 'content-control-tags.docx');
const ZIP_DATE = new Date('2026-10-08T12:00:00Z');

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

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const sdt = (tag: string, inner: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
const group = (index: number, first: string, second: string) =>
  sdt(`group-${index}`, sdt(`case-${index}`, run(first)) + sdt(`fallback-${index}`, run(second)));

const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}">
  <w:body>
    <w:p>${run('Status: ')}${group(1, 'alpha', 'beta')}${run('. Then: ')}${group(2, 'gamma', 'delta')}${run('.')}</w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1296" w:right="1296" w:bottom="1296" w:left="1296" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;

export async function createContentControlTagsFixture(): Promise<Uint8Array> {
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
  const bytes = await createContentControlTagsFixture();
  fs.writeFileSync(OUT, bytes);
  console.log(`Created ${OUT}`);
}
