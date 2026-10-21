// A sequence of labels at one edge is one tag: it lays out as a single unit with the room each of
// its labels keeps, stands between the edge's two slots like any tag, and paints each label as a
// part of its own, with that label's names, mark and variant.

import { GlobalRegistrator } from '@happy-dom/global-registrator';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, type OoxmlPart } from '@docx-editor.dev/core/store';
import type {
  ContentControlTagDisplay,
  ContentControlTagLabel,
} from '../../contracts/editor-content-control-view.ts';
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

const BODY =
  '<w:p><w:r><w:t xml:space="preserve">Declara </w:t></w:r>' +
  '<w:sdt><w:sdtPr><w:tag w:val="item"/></w:sdtPr>' +
  '<w:sdtContent><w:r><w:t>locado</w:t></w:r></w:sdtContent></w:sdt>' +
  '<w:r><w:t xml:space="preserve"> fim</w:t></w:r></w:p>';

const CONDITION = { text: 'há locação ▸', tone: 'optional' } as const;
const NUMBER = { text: 'II – ', tone: 'item', variant: 'text' } as const;
const SYMBOL = { text: '; e', tone: 'item', variant: 'text' } as const;
const CLOSE = { text: '◂', tone: 'optional' } as const;

const SEQUENCES: ContentControlTagDisplay = {
  labelsOf: () => ({
    open: { variant: 'sequence', labels: [CONDITION, NUMBER] },
    close: { variant: 'sequence', labels: [SYMBOL, CLOSE] },
  }),
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

/** The width one label takes as the only tag of its edge. */
function widthAlone(label: ContentControlTagLabel): number {
  return tagSpansOf(layoutOf({ labelsOf: () => ({ open: label }) }))[0]!.box.width;
}

describe('a sequence of labels at one edge', () => {
  test('lays out as one tag per edge, its labels one after the other', () => {
    const tags = tagSpansOf(layoutOf(SEQUENCES));

    expect(tags.map((span) => [span.contentControlTag!.edge, span.text])).toEqual([
      ['open', 'há locação ▸II – '],
      ['close', '; e◂'],
    ]);
  });

  test('takes the room each of its labels takes on its own', () => {
    const [open, close] = tagSpansOf(layoutOf(SEQUENCES));

    expect(open!.box.width).toBeCloseTo(widthAlone(CONDITION) + widthAlone(NUMBER), 6);
    expect(close!.box.width).toBeCloseTo(widthAlone(SYMBOL) + widthAlone(CLOSE), 6);
  });

  test('stands between the two slots of its edge, as any tag does', () => {
    const layout = layoutOf(SEQUENCES);
    const paragraphId = tagSpansOf(layout)[0]!.range.paragraphId;
    const sides = (offset: number) =>
      caretSlotsAt(layout, { paragraphId, offset }).map((slot) => [
        slot.left?.edge ?? null,
        slot.right?.edge ?? null,
      ]);

    expect(sides(8)).toEqual([
      [null, 'open'],
      ['open', null],
    ]);
    expect(sides(14)).toEqual([
      [null, 'close'],
      ['close', null],
    ]);
  });

  test('paints one tag per edge with a part per label, each with its own names and variant', () => {
    const container = document.createElement('div');
    paintSemanticLayout(container, layoutOf(SEQUENCES), { scale: 1, ariaHidden: false });
    const tags = [...container.querySelectorAll<HTMLElement>('[data-cc-tag-control]')];
    const parts = (tag: HTMLElement) =>
      [...tag.querySelectorAll<HTMLElement>('.docx-cc-tag')].map((part) => [
        part.textContent,
        part.dataset.pieceTone ?? null,
        part.dataset.ccTagVariant ?? null,
      ]);

    expect(tags.map((tag) => [tag.dataset.ccTagEdge, tag.dataset.ccTagVariant])).toEqual([
      ['open', 'sequence'],
      ['close', 'sequence'],
    ]);
    expect(tags.map(parts)).toEqual([
      [
        ['há locação ▸', 'optional', null],
        ['II – ', 'item', 'text'],
      ],
      [
        ['; e', 'item', 'text'],
        ['◂', 'optional', null],
      ],
    ]);
  });

  test('a chip part draws its mark in its own room; a text part takes no mark', () => {
    const marked = { ...CONDITION, mark: 'optional' } as const;
    const display: ContentControlTagDisplay = {
      labelsOf: () => ({
        open: { variant: 'sequence', labels: [marked, { ...NUMBER, mark: 'ignored' }] },
      }),
    };
    const [open] = tagSpansOf(layoutOf(display));
    expect(open!.contentControlTag!.parts!.map((part) => part.mark ?? null)).toEqual([
      'optional',
      null,
    ]);
    expect(open!.box.width).toBeCloseTo(widthAlone(marked) + widthAlone(NUMBER), 6);
    const container = document.createElement('div');
    paintSemanticLayout(container, layoutOf(display), { scale: 1, ariaHidden: false });
    const slots = [...container.querySelectorAll<HTMLElement>('[data-piece-mark]')];
    const labelOf = (slot: HTMLElement) => slot.parentElement?.textContent?.replace(/\u00A0/g, ' ');
    expect(slots.map((slot) => [slot.dataset.pieceMark, labelOf(slot)])).toEqual([
      ['optional', 'há locação ▸'],
    ]);
  });

  test('an empty label is left out, and a sequence with nothing to draw draws no tag', () => {
    const one = tagSpansOf(
      layoutOf({
        labelsOf: () => ({ open: { variant: 'sequence', labels: [{ text: '' }, NUMBER] } }),
      })
    );
    const none = tagSpansOf(
      layoutOf({ labelsOf: () => ({ open: { variant: 'sequence', labels: [{ text: '' }] } }) })
    );

    expect(one.map((span) => span.text)).toEqual(['II – ']);
    expect(none).toEqual([]);
  });
});
