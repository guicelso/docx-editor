// The tags of a BLOCK-level control: the open tag at the start of its first paragraph, the close tag
// at the end of its last, outside the inline tags there, with the face of the text beside them.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, type OoxmlPart } from '../../store/package/ooxml-tree.ts';
import { createFixedMeasurer, layoutSemanticDocument } from '../semantic-layout.ts';
import type { SemanticLayout, StyleSpanRecord } from '../semantic-records.ts';
import type { ContentControlTagDisplay } from '../content-control-tags.ts';
import { caretSlotsAt } from '../content-control-tag-slots.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function documentOf(bodyXml: string): OoxmlPart {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body>${bodyXml}</w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'app/xml' }
  );
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

const run = (text: string, rPr = '') => `<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const paragraph = (inner: string) => `<w:p>${inner}</w:p>`;
const inline = (tag: string, inner: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
const block = (tag: string, inner: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w:richText/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

const TAGS: ContentControlTagDisplay = {
  token: 'chips',
  labelsOf: ({ tag }) => {
    if (tag === 'only-open') return { open: { text: '[O' } };
    return tag === undefined ? null : { open: { text: `[${tag}` }, close: { text: `${tag}]` } };
  },
};

const measurer = createFixedMeasurer(6, 14);
const layoutOf = (body: string): SemanticLayout =>
  layoutSemanticDocument(documentOf(body), 0, { measurer, contentControlTags: TAGS });

/** Each paragraph's painted text, chips included, in document order. */
function paintedParagraphs(layout: SemanticLayout): string[] {
  const byParagraph = new Map<string, string>();
  for (const span of spansOf(layout)) {
    const id = span.range.paragraphId;
    byParagraph.set(id, (byParagraph.get(id) ?? '') + span.text.replace(/ /g, ' '));
  }
  return [...byParagraph.values()];
}

function spansOf(layout: SemanticLayout): (StyleSpanRecord & { pageIndex: number })[] {
  const spans: (StyleSpanRecord & { pageIndex: number })[] = [];
  layout.pages.forEach((page, pageIndex) => {
    for (const fragment of page.fragments)
      for (const line of (fragment as { lines?: { spans: StyleSpanRecord[] }[] }).lines ?? [])
        spans.push(...line.spans.map((span) => ({ ...span, pageIndex })));
  });
  return spans;
}

describe('block-control tags', () => {
  test('open at the first paragraph and close at the last, outside the inline tags', () => {
    const layout = layoutOf(
      paragraph(run('a')) +
        block('B', paragraph(inline('i', run('x'))) + paragraph(run('m')) + paragraph(run('y')))
    );

    expect(paintedParagraphs(layout)).toEqual(['a', '[B[ixi]', 'm', 'yB]']);
  });

  test('nested block controls open outer first and close inner first', () => {
    const layout = layoutOf(block('L', block('I', paragraph(run('x')))));

    expect(paintedParagraphs(layout)).toEqual(['[L[IxI]L]']);
  });

  test('a side the host leaves out has no tag', () => {
    const layout = layoutOf(block('only-open', paragraph(run('x'))));

    expect(paintedParagraphs(layout)).toEqual(['[Ox']);
  });

  test('a tag takes the face and size of the text beside it', () => {
    const layout = layoutOf(block('B', paragraph(run('x', '<w:rPr><w:sz w:val="40"/></w:rPr>'))));
    const tags = spansOf(layout).filter((span) => span.contentControlTag);

    expect(tags.map((span) => span.style.fontSizePt)).toEqual([20, 20]);
  });

  test('a control cut by a page has a tag at each end and none at the cut', () => {
    const body = Array.from({ length: 90 }, (_, index) => paragraph(run(`p${index}`))).join('');
    const layout = layoutOf(block('B', body));
    const tags = spansOf(layout).filter((span) => span.contentControlTag);
    const boundary = layout.contentControls?.find((control) => control.tag === 'B');

    expect(layout.pages.length).toBeGreaterThan(1);
    expect(tags.map((span) => [span.contentControlTag!.edge, span.pageIndex])).toEqual([
      ['open', 0],
      ['close', layout.pages.length - 1],
    ]);
    expect(boundary?.fragments.length).toBe(layout.pages.length);
  });

  test('the tag stands between two caret slots: outside on its left, inside on its right', () => {
    const layout = layoutOf(paragraph(run('a')) + block('B', paragraph(run('x'))));
    const first = spansOf(layout).find((span) => span.contentControlTag)!.range.paragraphId;
    const slots = caretSlotsAt(layout, { paragraphId: first, offset: 0 });

    expect(slots.map((slot) => [slot.left?.edge ?? null, slot.right?.edge ?? null])).toEqual([
      [null, 'open'],
      ['open', null],
    ]);
  });
});
