// An edge tag draws nothing and takes no room, and still stands as a tag: with a label on one side
// of a control and an edge on the other, the place just inside the control is a slot of its own on
// both sides, which is what lets typing at that side land in the control.

import { GlobalRegistrator } from '@happy-dom/global-registrator';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, type OoxmlPart } from '@docx-editor.dev/core/store';
import type { ContentControlTagDisplay } from '../../contracts/editor-content-control-view.ts';
import { caretSlotsAt } from '../../layout/content-control-tag-slots.ts';
import { createFixedMeasurer, layoutSemanticDocument } from '../../layout/semantic-layout.ts';
import type { StyleSpanRecord } from '../../layout/semantic-records.ts';
import { paintSemanticLayout } from '../semantic-paint.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const measurer = createFixedMeasurer(6, 14);

function load(body: string): OoxmlPart {
  const result = readOoxmlPart(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`, {
    name: '/word/document.xml',
    contentType: 'app/xml',
  });
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

// "Editor " + item{"4.1"} + " fim": the item ends in an edge, as the last item of a list does.
const BODY =
  '<w:p><w:r><w:t xml:space="preserve">Editor </w:t></w:r>' +
  '<w:sdt><w:sdtPr><w:tag w:val="item"/></w:sdtPr>' +
  '<w:sdtContent><w:r><w:t>4.1</w:t></w:r></w:sdtContent></w:sdt>' +
  '<w:r><w:t xml:space="preserve"> fim</w:t></w:r></w:p>';

const EDGES: ContentControlTagDisplay = {
  labelsOf: () => ({ open: { variant: 'edge' }, close: { variant: 'edge' } }),
};

function layoutOf(display: ContentControlTagDisplay) {
  return layoutSemanticDocument(load(BODY), 1, { measurer, contentControlView: { tags: display } });
}

function tagSpansOf(layout: ReturnType<typeof layoutSemanticDocument>): StyleSpanRecord[] {
  return layout.pages
    .flatMap((page) => page.fragments)
    .flatMap((fragment) => (fragment as { lines?: { spans: StyleSpanRecord[] }[] }).lines ?? [])
    .flatMap((line) => line.spans)
    .filter((span) => span.contentControlTag);
}

describe('an edge tag', () => {
  test('stands at each edge with no text and no room', () => {
    const tags = tagSpansOf(layoutOf(EDGES));

    expect(
      tags.map((span) => [span.contentControlTag!.edge, span.contentControlTag!.variant])
    ).toEqual([
      ['open', 'edge'],
      ['close', 'edge'],
    ]);
    expect(tags.map((span) => [span.text, span.box.width])).toEqual([
      ['', 0],
      ['', 0],
    ]);
  });

  test('makes the place just inside the control a slot of its own', () => {
    const layout = layoutOf(EDGES);
    const paragraphId = tagSpansOf(layout)[0]!.range.paragraphId;
    const sides = (offset: number) =>
      caretSlotsAt(layout, { paragraphId, offset }).map((slot) => [
        slot.left?.edge ?? null,
        slot.right?.edge ?? null,
      ]);

    expect(sides(7)).toEqual([
      [null, 'open'],
      ['open', null],
    ]);
    expect(sides(10)).toEqual([
      [null, 'close'],
      ['close', null],
    ]);
  });

  test('paints nothing a reader would see, and says what it is to the stylesheet', () => {
    const container = document.createElement('div');
    paintSemanticLayout(container, layoutOf(EDGES), { scale: 1, ariaHidden: false });
    const tags = [...container.querySelectorAll<HTMLElement>('[data-cc-tag-control]')];

    expect(tags.map((tag) => [tag.dataset.ccTagVariant, tag.textContent])).toEqual([
      ['edge', ''],
      ['edge', ''],
    ]);
    expect(tags.map((tag) => parseFloat(tag.style.width))).toEqual([0, 0]);
  });

  test('an empty label that is not an edge still draws no tag', () => {
    const tags = tagSpansOf(layoutOf({ labelsOf: () => ({ close: { text: '' } }) }));

    expect(tags).toEqual([]);
  });
});
