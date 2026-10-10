// View-only start and end tags for inline content controls — Word's "Design Mode" tags.
//
// A tag is a LAYOUT-OWNED piece: it measures, breaks and paints like text, and covers a
// ZERO-WIDTH model range at the control's edge, exactly as a `w:ptab` does. Nothing is written
// to the document, so every offset after a tag still agrees with the store.

import type { OoxmlElement } from '@docx-editor.dev/core/store';
import { contentControlPropertiesOf, propertyVal } from './content-control-properties.ts';
import type { FieldAwarePiece } from './field-pieces.ts';
import type { ResolvedRunStyle } from './run-style.ts';
import type { BlockControlEdges } from '../store/store/block-control-edges.ts';

/** The text a tag measures and paints: one unbreakable unit, its spaces made non-breaking. */
export function contentControlTagText(label: ContentControlTagLabel): string {
  return label.text.replace(/ /g, '\u00A0');
}

/** The transparent gap that keeps two neighbouring chips apart, inside each chip's advance. */
const CHIP_GAP_PT = 1.2;
/** The padding between the chip's fill and its label, in ems of the chip's own size. */
const CHIP_PAD_EM = 0.35;

/**
 * The room a chip keeps on EACH side of its label, in points: the gap between neighbours and the
 * padding inside the fill. Geometry, not characters — layout adds it to the measured label and
 * paint draws exactly it, so the label is centred in the space reserved for it whatever font
 * paints it.
 */
export function contentControlTagInsetsPt(style: Pick<ResolvedRunStyle, 'fontSizePt'>): {
  readonly gapPt: number;
  readonly padPt: number;
} {
  return { gapPt: CHIP_GAP_PT, padPt: style.fontSizePt * CHIP_PAD_EM };
}

/** What a chip adds to its label's measured width: the room on both sides. */
export function contentControlTagChromePt(style: Pick<ResolvedRunStyle, 'fontSizePt'>): number {
  const { gapPt, padPt } = contentControlTagInsetsPt(style);
  return 2 * (gapPt + padPt);
}

/** Which edge of the control a tag stands at. @public */
export type ContentControlTagEdge = 'open' | 'close';

/**
 * What a host draws at one edge. @public
 *
 * `tone` is an opaque name the chip publishes as `data-cc-tag-tone`, so the host's stylesheet
 * colours it with its own theme. Letters, digits and `-`, starting with a letter, at most 32;
 * anything else draws the neutral chip.
 */
export interface ContentControlTagLabel {
  readonly text: string;
  readonly tone?: string;
}

/** The control a host labels, as the file states it. `tag` is untrusted file data. @public */
export interface ContentControlTagSubject {
  readonly controlId: string;
  readonly tag: string | undefined;
}

/**
 * The host's answer for every inline control: the two labels, or null to draw none. @public
 *
 * Must be pure and cheap — it runs once per control per layout pass. When its answers change
 * the host re-installs it, which invalidates every cached page, as toggling field codes does.
 */
export interface ContentControlTagDisplay {
  /**
   * Names this labeling for the layout caches: change it whenever `labelsOf` would answer
   * differently, or cached paragraphs keep the old tags.
   */
  readonly token: string;
  readonly labelsOf: (
    control: ContentControlTagSubject
  ) => { readonly open?: ContentControlTagLabel; readonly close?: ContentControlTagLabel } | null;
}

/** Whether a tag's control sits inside a paragraph or holds whole paragraphs. @public */
export type ContentControlTagLevel = 'inline' | 'block';

/** What a tag piece, and every span cut from it, carries to paint and to the hit test. @public */
export interface ContentControlTagMark {
  readonly controlId: string;
  readonly edge: ContentControlTagEdge;
  readonly level: ContentControlTagLevel;
  /** The host's tone, already checked against {@link contentControlTagToneOf}. */
  readonly tone?: string;
}

const TONE = /^[A-Za-z][A-Za-z0-9-]{0,31}$/;

/** The tone a chip may publish, or undefined: the host's string lands in an attribute. */
export function contentControlTagToneOf(label: ContentControlTagLabel): string | undefined {
  return label.tone !== undefined && TONE.test(label.tone) ? label.tone : undefined;
}

/**
 * The piece a tag lays out as: its text over a ZERO-WIDTH range at the control's edge, so it
 * measures and paints like a word and never moves an offset. None for an empty label.
 */
export function contentControlTagPiece(
  mark: Omit<ContentControlTagMark, 'tone'>,
  label: ContentControlTagLabel,
  run: ResolvedRunStyle,
  offset: number
): FieldAwarePiece | null {
  if (label.text.length === 0) return null;
  const tone = contentControlTagToneOf(label);
  return {
    // A tag is one unit: line breaking must never open a line inside it.
    text: contentControlTagText(label),
    props: [],
    style: contentControlTagStyle(run),
    start: offset,
    end: offset,
    projected: true,
    contentControlTag: tone === undefined ? mark : { ...mark, tone },
  };
}

export function contentControlTagSubjectOf(control: OoxmlElement): ContentControlTagSubject {
  return { controlId: control.id, tag: propertyVal(contentControlPropertiesOf(control), 'tag') };
}

/**
 * Gives every tag the face and size of the text beside it: an open tag the text after it, a
 * close tag the text before it, and the other side when that side has none. A tag derived from
 * the paragraph's run properties alone is smaller than text formatted directly, and its chip
 * then sits short of the band a selection or a boundary draws over the line. Two passes, so a
 * paragraph of many empty controls stays linear.
 */
export function fitContentControlTagsToText(pieces: FieldAwarePiece[]): void {
  const before: (FieldAwarePiece | undefined)[] = [];
  let last: FieldAwarePiece | undefined;
  for (const piece of pieces) {
    before.push(last);
    if (isText(piece)) last = piece;
  }
  let next: FieldAwarePiece | undefined;
  for (let index = pieces.length - 1; index >= 0; index -= 1) {
    const piece = pieces[index]!;
    if (isText(piece)) next = piece;
    if (!piece.contentControlTag) continue;
    const inward = piece.contentControlTag.edge === 'open' ? next : before[index];
    const text = inward ?? (piece.contentControlTag.edge === 'open' ? before[index] : next);
    if (text) pieces[index] = { ...piece, style: contentControlTagStyle(text.style) };
  }
}

function isText(piece: FieldAwarePiece): boolean {
  return (
    piece.text.length > 0 &&
    !piece.contentControlTag &&
    !piece.inlineDrawing &&
    !piece.equation &&
    !piece.anchoredAtom &&
    !piece.noteSeparator
  );
}

/**
 * The style a tag measures with, derived from the run it sits in.
 *
 * The tag keeps the run's face and size, so its chip has the same band as the text beside it:
 * a selection or a control's boundary drawn over the line covers the chip exactly. Everything
 * else that would change the glyphs of the surrounding prose is reset — a tag inside a bold,
 * struck-through, all-caps run must not inherit any of it. Colour and fill are not run
 * formatting: the chip is painted by `semantic-paint` from the stylesheet, by tone.
 */
export function contentControlTagStyle(base: ResolvedRunStyle): ResolvedRunStyle {
  return {
    ...base,
    color: null,
    bold: false,
    italic: false,
    underline: null,
    strike: false,
    doubleStrike: false,
    highlight: null,
    shading: null,
    verticalAlign: 'baseline',
    baselineShiftPt: 0,
    caps: false,
    smallCaps: false,
    characterSpacingPt: 0,
    horizontalScalePercent: 100,
    hidden: false,
  };
}

/** What a lane forwards to draw tags: the host's labels and the block edges of the part. */
export interface ContentControlTagFlow {
  readonly contentControlTags?: ContentControlTagDisplay;
  readonly blockControlEdges?: ReadonlyMap<string, BlockControlEdges>;
}

/**
 * The tag inputs a lane hands the next one. The two travel together: block edges mean nothing
 * without labels to draw at them, and labels without them would draw no block tag.
 */
export function contentControlTagFlow(inputs: ContentControlTagFlow): ContentControlTagFlow {
  if (!inputs.contentControlTags) return {};
  return inputs.blockControlEdges
    ? { contentControlTags: inputs.contentControlTags, blockControlEdges: inputs.blockControlEdges }
    : { contentControlTags: inputs.contentControlTags };
}
