// A content-control tag is painted as a chip the HOST colours: the engine draws the shape,
// publishes which control, which edge and which tone, and writes no colour of its own.

import { GlobalRegistrator } from '@happy-dom/global-registrator';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, type OoxmlPart } from '@docx-editor.dev/core/store';
import {
  contentControlTagInsetsPt,
  type ContentControlTagDisplay,
} from '../../layout/content-control-tags.ts';
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
  '<w:p><w:r><w:t xml:space="preserve">CPF </w:t></w:r>' +
  '<w:sdt><w:sdtPr><w:tag w:val="span:optional:1"/></w:sdtPr>' +
  '<w:sdtContent><w:r><w:t>RG</w:t></w:r></w:sdtContent></w:sdt></w:p>';

function chips(display: ContentControlTagDisplay): HTMLElement[] {
  const layout = layoutSemanticDocument(load(BODY), 1, { measurer, contentControlTags: display });
  const container = document.createElement('div');
  paintSemanticLayout(container, layout, { scale: 1, ariaHidden: false });
  return [...container.querySelectorAll<HTMLElement>('[data-cc-tag-control]')];
}

describe('painting a content-control tag', () => {
  test('publishes the control, the edge and the host tone, and is a chip by class', () => {
    const [open, close] = chips({
      token: 't',
      labelsOf: () => ({ open: { text: 'Se', tone: 'optional' }, close: { text: '◂' } }),
    });
    expect(open?.classList.contains('docx-cc-tag')).toBe(true);
    expect(open?.dataset.ccTagEdge).toBe('open');
    expect(open?.dataset.ccTagTone).toBe('optional');
    expect(close?.dataset.ccTagEdge).toBe('close');
    expect(close?.dataset.ccTagTone).toBeUndefined();
    expect(open?.dataset.ccTagControl).toBe(close?.dataset.ccTagControl);
  });

  test('a tone that is not a plain name never reaches the attribute', () => {
    const [open] = chips({
      token: 't',
      labelsOf: () => ({ open: { text: 'Se', tone: 'x" onclick="alert(1)' } }),
    });
    expect(open?.dataset.ccTagTone).toBeUndefined();
  });

  test('writes no colour, fill or shadow of its own: those are the stylesheet’s', () => {
    const [open] = chips({
      token: 't',
      labelsOf: () => ({ open: { text: 'Se', tone: 'optional' } }),
    });
    expect(open?.style.color).toBe('');
    expect(open?.style.backgroundColor).toBe('');
    expect(open?.style.boxShadow).toBe('');
  });

  test('is furniture: hidden from assistive technology and not editable', () => {
    const [open] = chips({ token: 't', labelsOf: () => ({ open: { text: 'Se' } }) });
    expect(open?.getAttribute('aria-hidden')).toBe('true');
    expect(open?.contentEditable).toBe('false');
    expect(open?.dataset.paragraphId).toBeUndefined();
  });

  test('keeps the advance layout reserved, so nothing after it moves', () => {
    const { chip, record } = paintedChip({ open: { text: 'Se' } });
    const advance =
      parseFloat(chip.style.marginLeft) +
      parseFloat(chip.style.width) +
      parseFloat(chip.style.marginRight);
    expect(advance).toBeCloseTo(record.box.width * 2, 4);
  });

  test('draws the room layout measured: the same gap and padding on both sides of the label', () => {
    const { chip, record } = paintedChip({ close: { text: '◂' } });
    const { gapPt, padPt } = contentControlTagInsetsPt(record.style);
    expect(chip.style.marginLeft).toBe(chip.style.marginRight);
    expect(chip.style.paddingLeft).toBe(chip.style.paddingRight);
    expect(parseFloat(chip.style.marginLeft)).toBeCloseTo(gapPt * 2, 4);
    expect(parseFloat(chip.style.paddingLeft)).toBeCloseTo(padPt * 2, 4);
    // The label alone fills what is left — it never overflows into one side.
    const labelWidth = parseFloat(chip.style.width) - 2 * padPt * 2;
    expect(labelWidth).toBeCloseTo(measurer.measure('◂', record.style) * 2, 4);
  });

  test('keeps the gap outside the fill, so no border clips its corners', () => {
    const { chip } = paintedChip({ open: { text: 'Se' } });
    expect(chip.style.borderLeftWidth).toBe('');
    expect(chip.style.borderRightWidth).toBe('');
  });

  test('takes the band of the run it sits in, as the text beside it does', () => {
    const { chip, container } = paintedChip({ open: { text: 'Se' } });
    const text = container.querySelector<HTMLElement>('.layout-run[data-start]')!;
    expect(chip.style.height).toBe(text.style.height);
    expect(chip.style.lineHeight).toBe(text.style.lineHeight);
  });

  test('a justify gap before a chip comes on top of its own gap', () => {
    // A scaled run cannot stretch its trailing space, so the justify gap after it is a margin.
    const scaled = '<w:rPr><w:w w:val="90"/></w:rPr>';
    const justified =
      '<w:p><w:pPr><w:jc w:val="both"/></w:pPr>' +
      `<w:r>${scaled}<w:t xml:space="preserve">um dois tres quatro cinco seis sete oito nove dez </w:t></w:r>` +
      '<w:sdt><w:sdtPr><w:tag w:val="span:optional:1"/></w:sdtPr>' +
      '<w:sdtContent><w:r><w:t xml:space="preserve">onze doze treze quatorze quinze dezesseis</w:t></w:r></w:sdtContent></w:sdt>' +
      '<w:r><w:t xml:space="preserve"> dezessete dezoito dezenove vinte vinte e um vinte e dois</w:t></w:r></w:p>';
    const layout = layoutSemanticDocument(load(justified), 1, {
      measurer,
      contentControlTags: { token: 't', labelsOf: () => ({ open: { text: 'Se' } }) },
    });
    const container = document.createElement('div');
    paintSemanticLayout(container, layout, { scale: 1, ariaHidden: false });
    const chip = container.querySelector<HTMLElement>('[data-cc-tag-control]')!;
    const line = layout.pages
      .flatMap((page) => page.fragments)
      .flatMap((fragment) => (fragment as { lines?: { spans: StyleSpanRecord[] }[] }).lines ?? [])
      .find((candidate) => candidate.spans.some((span) => span.contentControlTag))!;
    const index = line.spans.findIndex((span) => span.contentControlTag);
    const [previous, record] = [line.spans[index - 1]!, line.spans[index]!];
    const justifyGap = record.box.x - (previous.box.x + previous.box.width);
    expect(justifyGap).toBeGreaterThan(0.001);
    const { gapPt } = contentControlTagInsetsPt(record.style);
    expect(parseFloat(chip.style.marginLeft)).toBeCloseTo(justifyGap + gapPt, 4);
  });
});

function tagRecordOf(layout: ReturnType<typeof layoutSemanticDocument>): StyleSpanRecord {
  return layout.pages
    .flatMap((page) => page.fragments)
    .flatMap((fragment) => (fragment as { lines?: { spans: StyleSpanRecord[] }[] }).lines ?? [])
    .flatMap((line) => line.spans)
    .find((span) => span.contentControlTag)!;
}

function paintedChip(labels: ReturnType<ContentControlTagDisplay['labelsOf']>): {
  chip: HTMLElement;
  record: StyleSpanRecord;
  container: HTMLElement;
} {
  const layout = layoutSemanticDocument(load(BODY), 1, {
    measurer,
    contentControlTags: { token: 't', labelsOf: () => labels },
  });
  const container = document.createElement('div');
  paintSemanticLayout(container, layout, { scale: 2, ariaHidden: false });
  const chip = container.querySelector<HTMLElement>('[data-cc-tag-control]')!;
  return { chip, record: tagRecordOf(layout), container };
}
