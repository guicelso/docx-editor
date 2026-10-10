// A tag label is a chip or text: text takes no room beside its label and draws no fill, and a side
// a host leaves out draws nothing. Either way the tag is a zero-width piece with the same slots.

import { GlobalRegistrator } from '@happy-dom/global-registrator';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, type OoxmlPart } from '@docx-editor.dev/core/store';
import type { ContentControlTagDisplay } from '../../contracts/editor-content-control-view.ts';
import { contentControlTagChromePt } from '../../layout/content-control-tags.ts';
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

// "Editor " + item{"4.1"} + " fim": the list item has no pill, its separator reads as text.
const BODY =
  '<w:p><w:r><w:t xml:space="preserve">Editor </w:t></w:r>' +
  '<w:sdt><w:sdtPr><w:tag w:val="item"/></w:sdtPr>' +
  '<w:sdtContent><w:r><w:t>4.1</w:t></w:r></w:sdtContent></w:sdt>' +
  '<w:r><w:t xml:space="preserve"> fim</w:t></w:r></w:p>';

const CLOSE_ONLY: ContentControlTagDisplay = {
  labelsOf: () => ({ close: { text: ', ', tone: 'list', variant: 'text' } }),
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

function paintedTags(display: ContentControlTagDisplay): HTMLElement[] {
  const container = document.createElement('div');
  paintSemanticLayout(container, layoutOf(display), { scale: 1, ariaHidden: false });
  return [...container.querySelectorAll<HTMLElement>('[data-cc-tag-control]')];
}

describe('the variants of a tag label', () => {
  test('a text tag measures its label alone; a chip adds its room', () => {
    const [text] = tagSpansOf(layoutOf(CLOSE_ONLY));
    expect(text!.box.width).toBeCloseTo(measurer.measure(text!.text, text!.style), 6);
    const [chip] = tagSpansOf(layoutOf({ labelsOf: () => ({ close: { text: ', ' } }) }));
    expect(contentControlTagChromePt(chip!.style, chip!.contentControlTag!)).toBeGreaterThan(0);
    expect(chip!.box.width).toBeCloseTo(
      measurer.measure(chip!.text, chip!.style) +
        contentControlTagChromePt(chip!.style, chip!.contentControlTag!),
      6
    );
  });

  test('a text tag paints with no room around its label and says so to the stylesheet', () => {
    const [text] = paintedTags(CLOSE_ONLY);
    expect(text!.dataset.ccTagVariant).toBe('text');
    expect(text!.dataset.ccTagTone).toBe('list');
    expect(parseFloat(text!.style.marginLeft)).toBe(0);
    expect(parseFloat(text!.style.paddingLeft)).toBe(0);
    const [chip] = paintedTags({ labelsOf: () => ({ close: { text: ', ' } }) });
    expect(chip!.dataset.ccTagVariant).toBeUndefined();
  });

  test('the side a host leaves out draws no tag', () => {
    const tags = tagSpansOf(layoutOf(CLOSE_ONLY));
    expect(tags.map((span) => span.contentControlTag!.edge)).toEqual(['close']);
  });

  test('a text tag is a zero-width piece with a slot on each side, as a chip is', () => {
    const layout = layoutOf(CLOSE_ONLY);
    const [tag] = tagSpansOf(layout);
    expect([tag!.range.start, tag!.range.end]).toEqual([10, 10]);
    const slots = caretSlotsAt(layout, { paragraphId: tag!.range.paragraphId, offset: 10 });
    expect(slots.map((slot) => [slot.left?.edge ?? null, slot.right?.edge ?? null])).toEqual([
      [null, 'close'],
      ['close', null],
    ]);
  });
});
