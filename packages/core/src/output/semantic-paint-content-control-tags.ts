// View-only content-control tags as the page paints them: a chip at a control's edge, and the parts of
// a sequence of labels. Split out of `semantic-paint.ts` so the tag's geometry — the room layout
// reserved for it — lives next to the one function that measures it.

import { contentControlTagInsetsPt } from '../layout/content-control-tags.ts';
import type { StyleSpanRecord } from '@docx-editor.dev/core/layout';

/**
 * A view-only content-control tag drawn as its own chip.
 *
 * The chip keeps EXACTLY the advance layout reserved — the caret and every following glyph are
 * placed from that number. The gap between neighbouring chips is a margin on each side, outside
 * the fill: a transparent border there clipped the fill's corners square. Its height is the
 * run's band, like any span's, so a selection or a boundary drawn over the line covers it
 * exactly. Colour, fill and outline belong to the stylesheet: `.docx-cc-tag` draws a neutral chip
 * from the `--doc-*` tokens, `data-cc-tag-variant="text"` draws no fill, and a host restyles it by
 * the `data-cc-tag-tone` it chose. The data attributes are also the host's click target.
 */
export function paintContentControlTagChip(
  element: HTMLElement,
  span: StyleSpanRecord,
  scale: number
): void {
  const tag = span.contentControlTag!;
  const css = element.style;
  element.classList.add('docx-cc-tag');
  const { gapPt, padPt } = contentControlTagInsetsPt(span.style, tag);
  css.width = `${(span.box.width - 2 * gapPt) * scale}px`;
  css.marginLeft = `${gapPt * scale}px`;
  css.marginRight = `${gapPt * scale}px`;
  css.paddingLeft = `${padPt * scale}px`;
  css.paddingRight = `${padPt * scale}px`;
  css.textAlign = 'center';
  // No `overflow: hidden`: on an inline-block it moves the baseline to the bottom edge
  // (CSS 2.1 §10.8.1) and lifts the chip off the line. The label fits by construction —
  // layout measured this very text for the advance.
  css.whiteSpace = 'nowrap';
  element.dataset.ccTagControl = tag.controlId;
  element.dataset.ccTagEdge = tag.edge;
  if (tag.tone !== undefined) element.dataset.ccTagTone = tag.tone;
  if (tag.variant !== undefined) element.dataset.ccTagVariant = tag.variant;
}

/**
 * A sequence paints each of its labels as a part of its own, with the room that label keeps and
 * the tone and variant it chose, so the stylesheet colours each part as it would a tag alone. The
 * parts share the tag's band, and together fill the advance layout measured for the whole text.
 */
export function mountContentControlTagParts(
  document: Document,
  host: HTMLElement,
  span: StyleSpanRecord,
  scale: number
): void {
  let from = 0;
  for (const part of span.contentControlTag!.parts!) {
    const element = document.createElement('span');
    const { gapPt, padPt } = contentControlTagInsetsPt(span.style, part);
    element.classList.add('docx-cc-tag');
    element.style.display = 'inline-block';
    element.style.height = '100%';
    element.style.marginLeft = `${gapPt * scale}px`;
    element.style.marginRight = `${gapPt * scale}px`;
    element.style.paddingLeft = `${padPt * scale}px`;
    element.style.paddingRight = `${padPt * scale}px`;
    if (part.tone !== undefined) element.dataset.ccTagTone = part.tone;
    if (part.variant !== undefined) element.dataset.ccTagVariant = part.variant;
    element.textContent = span.text.slice(from, from + part.length); // SAFE: textContent, never innerHTML
    from += part.length;
    host.append(element);
  }
}

/** The margin a chip keeps on its left, inside its advance; none for any other span. */
export function chipGapPt(span: StyleSpanRecord): number {
  return span.contentControlTag
    ? contentControlTagInsetsPt(span.style, span.contentControlTag).gapPt
    : 0;
}
