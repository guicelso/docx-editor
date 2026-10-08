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
    const layout = layoutSemanticDocument(load(BODY), 1, {
      measurer,
      contentControlTags: { token: 't', labelsOf: () => ({ open: { text: 'Se' } }) },
    });
    const container = document.createElement('div');
    paintSemanticLayout(container, layout, { scale: 2, ariaHidden: false });
    const chip = container.querySelector<HTMLElement>('[data-cc-tag-control]')!;
    const record = layout.pages
      .flatMap((page) => page.fragments)
      .flatMap(
        (fragment) =>
          (
            fragment as {
              lines?: { spans: { contentControlTag?: unknown; box: { width: number } }[] }[];
            }
          ).lines ?? []
      )
      .flatMap((line) => line.spans)
      .find((span) => span.contentControlTag)!;
    expect(parseFloat(chip.style.width)).toBeCloseTo(record.box.width * 2, 4);
  });

  test('draws the room layout measured: the same gap and padding on both sides of the label', () => {
    const layout = layoutSemanticDocument(load(BODY), 1, {
      measurer,
      contentControlTags: { token: 't', labelsOf: () => ({ close: { text: '◂' } }) },
    });
    const container = document.createElement('div');
    paintSemanticLayout(container, layout, { scale: 2, ariaHidden: false });
    const chip = container.querySelector<HTMLElement>('[data-cc-tag-control]')!;
    const record = layout.pages
      .flatMap((page) => page.fragments)
      .flatMap((fragment) => (fragment as { lines?: { spans: StyleSpanRecord[] }[] }).lines ?? [])
      .flatMap((line) => line.spans)
      .find((span) => span.contentControlTag)!;
    const { gapPt, padPt } = contentControlTagInsetsPt(record.style);
    expect(chip.style.borderLeftWidth).toBe(chip.style.borderRightWidth);
    expect(chip.style.paddingLeft).toBe(chip.style.paddingRight);
    expect(parseFloat(chip.style.borderLeftWidth)).toBeCloseTo(gapPt * 2, 4);
    expect(parseFloat(chip.style.paddingLeft)).toBeCloseTo(padPt * 2, 4);
    // The label alone fills what is left — it never overflows into one side.
    const labelWidth = parseFloat(chip.style.width) - 2 * (gapPt + padPt) * 2;
    expect(labelWidth).toBeCloseTo(measurer.measure('◂', record.style) * 2, 4);
  });
});
