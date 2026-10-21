// View-only start and end tags for inline content controls — Word's "Design Mode" tags.
//
// A tag is a LAYOUT-OWNED piece: it measures, breaks and paints like text, and covers a
// ZERO-WIDTH model range at the control's edge, exactly as a `w:ptab` does. Nothing is written
// to the document, so every offset after a tag still agrees with the store.

import type {
  ContentControlTagInk,
  ContentControlTagLabel,
} from '../contracts/editor-content-control-view.ts';
import type { FieldAwarePiece } from './field-pieces.ts';
import {
  checkedPieceNameOf,
  NO_PIECE_ROOM,
  pieceNameRecordOf,
  pieceRoomOf,
  pieceRoomWidthPt,
  type PieceNameRecord,
  type PieceRoom,
} from './piece-names.ts';
import type { ResolvedRunStyle } from './run-style.ts';

/** The text a tag measures and paints: one unbreakable unit, its spaces made non-breaking. */
export function contentControlTagText(label: ContentControlTagInk): string {
  return label.text.replace(/ /g, ' ');
}

/**
 * The room a tag, or one part of a sequence, keeps around its label: a chip's (with its mark), and
 * none for a `text`, an `edge` or a `sequence` tag, whose parts keep their own.
 */
export function contentControlTagRoomOf(
  style: Pick<ResolvedRunStyle, 'fontSizePt'>,
  mark: Pick<ContentControlTagMark, 'variant' | 'mark'>
): PieceRoom {
  if (mark.variant !== undefined) return NO_PIECE_ROOM;
  return pieceRoomOf(style, mark.mark);
}

/** What a tag adds to its label's measured width: the room of each part of a sequence. */
export function contentControlTagChromePt(
  style: Pick<ResolvedRunStyle, 'fontSizePt'>,
  mark: Pick<ContentControlTagMark, 'variant' | 'mark' | 'parts'>
): number {
  const parts = mark.parts ?? [mark];
  return parts.reduce(
    (sum, part) => sum + pieceRoomWidthPt(contentControlTagRoomOf(style, part)),
    0
  );
}

/** Which edge of the control a tag stands at. @public */
export type ContentControlTagEdge = 'open' | 'close';

/** Whether a tag's control sits inside a paragraph or holds whole paragraphs. @public */
export type ContentControlTagLevel = 'inline' | 'block';

/** What a tag piece, and every span cut from it, carries to paint and to the hit test. @public */
export interface ContentControlTagMark {
  readonly controlId: string;
  readonly edge: ContentControlTagEdge;
  readonly level: ContentControlTagLevel;
  /** The host's names, already checked. */
  readonly name?: PieceNameRecord;
  /** A chip's mark, already checked. */
  readonly mark?: string;
  /**
   * Present on a tag of a control the host made an atom: the tag is part of that unit, painted
   * with its control so it carries the unit's hover and selection.
   */
  readonly atom?: true;
  /**
   * Present for a tag drawn as text in the line, one that draws nothing, or a sequence of labels;
   * absent for a chip.
   */
  readonly variant?: 'text' | 'edge' | 'sequence';
  /** The labels of a sequence, in order, as many characters of the tag's text as each one paints. */
  readonly parts?: readonly ContentControlTagPart[];
}

/** One label of a sequence: how much of the tag's text it paints, its names, mark and variant. @public */
export interface ContentControlTagPart {
  readonly length: number;
  /** The label's names, already checked. */
  readonly name?: PieceNameRecord;
  /** A chip label's mark, already checked. */
  readonly mark?: string;
  /** Present for a label drawn as text in the line; absent for a chip. */
  readonly variant?: 'text';
}

/**
 * The piece a tag lays out as: its text over a ZERO-WIDTH range at the control's edge, so it
 * measures and paints like a word and never moves an offset. None for an empty label; an `edge`
 * is a piece with no text, which takes no room and still stands between its two slots.
 */
export function contentControlTagPiece(
  mark: Omit<ContentControlTagMark, 'name' | 'mark' | 'variant' | 'parts'>,
  label: ContentControlTagLabel,
  run: ResolvedRunStyle,
  offset: number
): FieldAwarePiece | null {
  const drawn = drawnTagOf(label);
  if (drawn === null) return null;
  return {
    // A tag is one unit: line breaking must never open a line inside it.
    text: drawn.text,
    props: [],
    style: contentControlTagStyle(run),
    start: offset,
    end: offset,
    projected: true,
    contentControlTag: { ...mark, ...drawn.mark },
  };
}

type DrawnInk = Pick<ContentControlTagMark, 'name' | 'mark'> & { readonly variant?: 'text' };

/** The text a label paints and what its mark says of it, or null for a label with nothing to draw. */
function drawnTagOf(label: ContentControlTagLabel): Readonly<{
  text: string;
  mark: Pick<ContentControlTagMark, 'name' | 'mark' | 'variant' | 'parts'>;
}> | null {
  if (label.variant === 'edge') return { text: '', mark: { variant: 'edge' } };
  if (label.variant === 'sequence') return drawnSequenceOf(label.labels);
  if (label.text.length === 0) return null;
  return { text: contentControlTagText(label), mark: drawnInkOf(label) };
}

function drawnSequenceOf(labels: readonly ContentControlTagInk[]): ReturnType<typeof drawnTagOf> {
  const drawn = labels.filter((label) => label.text.length > 0);
  if (drawn.length === 0) return null;
  const texts = drawn.map(contentControlTagText);
  const parts = drawn.map(
    (label, index): ContentControlTagPart => ({
      length: texts[index]!.length,
      ...drawnInkOf(label),
    })
  );
  return { text: texts.join(''), mark: { variant: 'sequence', parts } };
}

/** A label's checked names and mark; a mark only on a chip, which has the room for it. */
function drawnInkOf(label: ContentControlTagInk): DrawnInk {
  const name = pieceNameRecordOf(label);
  const mark = label.variant === 'text' ? undefined : checkedPieceNameOf(label.mark);
  return {
    ...(name === undefined ? {} : { name }),
    ...(mark === undefined ? {} : { mark }),
    ...(label.variant === 'text' ? { variant: label.variant } : {}),
  };
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
 * formatting: the chip is painted by `semantic-paint` from the stylesheet, by its names.
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
