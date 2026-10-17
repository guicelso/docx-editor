// Word's Web Layout (`ST_View` `web`): one sheet as wide as the host and as tall as its text,
// where breaks open no sheet, and what the view does not lay out is named rather than drawn.

import { describe, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import { readOoxmlPackage, type OoxmlPart } from '@docx-editor.dev/core/store';
import {
  createFixedMeasurer,
  createLayoutSession,
  layoutSemanticDocument,
  type SemanticLayout,
} from '../index.ts';
import { webLayoutViewRefusal } from '../web-layout-view.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const measurer = createFixedMeasurer(6, 14);

function bodyPart(body: string): OoxmlPart {
  const bytes = zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '</Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="${WP}"><w:body>${body}</w:body></w:document>`
    ),
  });
  const loaded = readOoxmlPackage(bytes);
  if (!loaded.ok) throw new Error(loaded.error.message);
  return loaded.package.parts.get(loaded.package.mainDocumentPart)!;
}

const paragraph = (text: string, properties = '') =>
  `<w:p>${properties ? `<w:pPr>${properties}</w:pPr>` : ''}<w:r><w:t>${text}</w:t></w:r></w:p>`;
const paragraphs = (count: number) =>
  Array.from({ length: count }, (_, index) => paragraph(`Line ${index}`)).join('');
const SECTION = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';

const web = (widthPt: number) => ({ kind: 'web' as const, widthPt });

/** Where the last block of the only sheet ends. */
const textBottom = (layout: SemanticLayout): number =>
  Math.max(...layout.pages[0]!.fragments.map((fragment) => fragment.box.y + fragment.box.height));

describe('the web layout view', () => {
  test('lays out one sheet as wide as the host and as tall as its text', () => {
    const part = bodyPart(paragraphs(80) + SECTION);

    const printed = layoutSemanticDocument(part, 1, { measurer });
    const laid = layoutSemanticDocument(part, 1, { measurer, layoutView: web(300) });

    expect(printed.pages.length).toBeGreaterThan(1);
    expect(laid.pages).toHaveLength(1);
    const [page] = laid.pages;
    const height = textBottom(laid);
    expect(height).toBeGreaterThan(792);
    expect(page!.box).toEqual({ x: 0, y: 0, width: 300, height });
    expect(page!.contentBox).toEqual({ x: 0, y: 0, width: 300, height });
    expect(page!.header).toBeUndefined();
    expect(page!.footer).toBeUndefined();
  });

  test('a page break, a column break and page-break-before open no sheet', () => {
    const part = bodyPart(
      paragraph('first') +
        '<w:p><w:r><w:t>before</w:t><w:br w:type="page"/><w:t>after</w:t></w:r></w:p>' +
        paragraph('kept', '<w:pageBreakBefore/>') +
        '<w:p><w:r><w:t>left</w:t><w:br w:type="column"/><w:t>right</w:t></w:r></w:p>' +
        SECTION
    );

    const laid = layoutSemanticDocument(part, 1, { measurer, layoutView: web(400) });

    expect(laid.pages).toHaveLength(1);
    const tops = laid.pages[0]!.fragments.map((fragment) => fragment.box.y);
    expect(tops).toEqual([...tops].sort((a, b) => a - b));
    expect(new Set(tops).size).toBe(tops.length);
  });

  test('several sections flow as one, at the host width rather than any section’s', () => {
    const part = bodyPart(
      paragraph('portrait') +
        '<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr></w:p>' +
        paragraph('landscape') +
        '<w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/></w:sectPr>'
    );

    const laid = layoutSemanticDocument(part, 1, { measurer, layoutView: web(250) });

    expect(laid.pages).toHaveLength(1);
    expect(laid.pages[0]!.box.width).toBe(250);
  });

  test('grows with the text across incremental passes, and leaves them for sheets', () => {
    const session = createLayoutSession();
    const before = layoutSemanticDocument(bodyPart(paragraphs(3) + SECTION), 1, {
      measurer,
      session,
      layoutView: web(300),
    });
    const after = layoutSemanticDocument(bodyPart(paragraphs(5) + SECTION), 2, {
      measurer,
      session,
      layoutView: web(300),
    });
    const printed = layoutSemanticDocument(bodyPart(paragraphs(5) + SECTION), 3, {
      measurer,
      session,
    });

    const line = before.pages[0]!.box.height / 3;
    expect(before.pages[0]!.box.height).toBe(textBottom(before));
    expect(after.pages[0]!.box.height).toBeCloseTo(5 * line, 6);
    expect(printed.pages[0]!.box.height).toBe(792);
  });
});

describe('what the web layout view refuses', () => {
  const refusalOf = (body: string) => webLayoutViewRefusal(bodyPart(body), undefined, 300);

  test('admits text, inline content and tables in flow', () => {
    expect(refusalOf(paragraphs(2) + SECTION)).toBeNull();
    expect(
      refusalOf(
        '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
          `<w:tr><w:tc>${paragraph('cell')}</w:tc></w:tr></w:tbl>` +
          SECTION
      )
    ).toBeNull();
  });

  test('names columns and a document grid, which are page setup', () => {
    expect(refusalOf(paragraph('a') + '<w:sectPr><w:cols w:num="2"/></w:sectPr>')).toBe('columns');
    expect(
      refusalOf(
        paragraph('a') + '<w:sectPr><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>'
      )
    ).toBe('line-grid');
  });

  test('names notes, floating pictures, floating tables and text frames', () => {
    expect(refusalOf('<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>' + SECTION)).toBe(
      'notes'
    );
    expect(
      refusalOf(
        '<w:p><w:r><w:drawing><wp:anchor simplePos="0" relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
          '<wp:simplePos x="0" y="0"/><wp:extent cx="100" cy="100"/></wp:anchor></w:drawing></w:r></w:p>' +
          SECTION
      )
    ).toBe('floating-drawing');
    expect(
      refusalOf(
        '<w:tbl><w:tblPr><w:tblpPr w:vertAnchor="page" w:tblpY="1440"/></w:tblPr>' +
          '<w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
          `<w:tr><w:tc>${paragraph('cell')}</w:tc></w:tr></w:tbl>` +
          paragraph('after') +
          SECTION
      )
    ).toBe('floating-table');
    expect(
      refusalOf(paragraph('framed', '<w:framePr w:w="2000" w:hAnchor="page"/>') + SECTION)
    ).toBe('text-frame');
  });
});
