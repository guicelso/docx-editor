// How a view-only content-control tag sits in a line: where a line may open beside it, and how a
// tag that no line can hold is cut to fit.
//
// A tag is chrome, not a letter of the word beside it. Glued to that word it made one unit with
// every tag and character that touched it, and a chain of tags wider than the line stood past the
// margin, because a word that opens its line is never moved again.

import type { FieldAwarePiece } from './field-pieces.ts';
import { segmentGraphemes } from './grapheme.ts';
import { pieceChromePt, pieceTextWidthPt } from './layout-owned-piece.ts';
import { displayText } from './run-style.ts';
import { styleForFontSlot } from './script-itemization.ts';
import type { TextMeasurer } from './semantic-records.ts';

/**
 * Whether a line may open between two pieces at a content-control tag: before an open tag and
 * after a close tag, as around an inline object. The tags of nested controls travel together (two
 * opens, or two closes), and so does punctuation that may not end a line (an opening bracket or
 * quote) or start one (a closing bracket or quote, a comma, a stop).
 */
export function opensAtTagSeam(
  previous: FieldAwarePiece | undefined,
  current: FieldAwarePiece
): boolean {
  if (previous === undefined) return false;
  const before = previous.contentControlTag?.edge;
  const after = current.contentControlTag?.edge;
  if (after === 'open' && before !== 'open') return !ENDS_WITH_OPENING_MARK.test(previous.text);
  if (before === 'close' && after !== 'close') return !STARTS_WITH_CLOSING_MARK.test(current.text);
  return false;
}

const ENDS_WITH_OPENING_MARK = /[\p{Ps}\p{Pi}]$/u;

const STARTS_WITH_CLOSING_MARK = /^[\p{Pe}\p{Pf},.;:!?…%]/u;

const ELLIPSIS = '…';

/**
 * The pieces with every unit that holds a tag cut to fit `measurePt`. A unit is what travels
 * together: the tags glued to each other and to the words they touch, up to a breaking space or a
 * tag seam. Its tags share the room left by its text — each one wider than a common cap keeps the
 * start of its longest label and ends in an ellipsis — so no tag gives up all its room to another.
 * Cutting a tag loses no text of the document, and the host still names the control in full. The
 * same pieces reach placement and the anchor probe.
 */
export function withContentControlTagsFitted(
  pieces: readonly FieldAwarePiece[],
  measurePt: number,
  measurer: TextMeasurer
): readonly FieldAwarePiece[] {
  if (!pieces.some((piece) => piece.contentControlTag)) return pieces;
  const fitted = [...pieces];
  for (const unit of gluedUnitsOf(pieces, measurer)) {
    const widths = unit.tags.map((index) =>
      widthOfPt(pieces[index]!, pieces[index]!.text, measurer)
    );
    const cap = sharedCapPt(widths, measurePt - unit.textWidthPt);
    unit.tags.forEach((index, at) => {
      if (widths[at]! > cap + 0.001) fitted[index] = fittedTag(pieces[index]!, cap, measurer);
    });
  }
  return fitted;
}

/** The tags of a unit, by index, and the width of the text that travels with them. */
interface GluedUnit {
  readonly tags: number[];
  textWidthPt: number;
}

function gluedUnitsOf(pieces: readonly FieldAwarePiece[], measurer: TextMeasurer): GluedUnit[] {
  const units: GluedUnit[] = [];
  let unit: GluedUnit = { tags: [], textWidthPt: 0 };
  const close = (): void => {
    if (unit.tags.length > 0) units.push(unit);
    unit = { tags: [], textWidthPt: 0 };
  };
  pieces.forEach((piece, index) => {
    if (opensAtTagSeam(pieces[index - 1], piece)) close();
    if (piece.contentControlTag) {
      unit.tags.push(index);
      return;
    }
    if (piece.pieceAtom) {
      unit.textWidthPt += widthOfPt(piece, piece.text, measurer);
      return;
    }
    const words = piece.text.split(/\s/u);
    unit.textWidthPt += widthOfPt(piece, words[0]!, measurer);
    if (words.length === 1) return;
    close();
    unit.textWidthPt += widthOfPt(piece, words.at(-1)!, measurer);
  });
  close();
  return units;
}

/** The largest width each tag may keep so that the tags, capped at it, fit `roomPt`. */
function sharedCapPt(widths: readonly number[], roomPt: number): number {
  const total = (cap: number): number =>
    widths.reduce((sum, width) => sum + Math.min(width, cap), 0);
  let low = 0;
  let high = widths.reduce((widest, width) => Math.max(widest, width), 0);
  if (total(high) <= roomPt + 0.001) return high;
  for (let step = 0; step < 32; step += 1) {
    const middle = (low + high) / 2;
    if (total(middle) <= roomPt) low = middle;
    else high = middle;
  }
  return low;
}

function widthOfPt(piece: FieldAwarePiece, text: string, measurer: TextMeasurer): number {
  const style = styleForFontSlot(piece.style, piece.fontSlot);
  return (
    pieceTextWidthPt(piece, text, style, (shown, face) =>
      measurer.measure(displayText(shown, face), face)
    ) + pieceChromePt(piece, style)
  );
}

function fittedTag(
  piece: FieldAwarePiece,
  measurePt: number,
  measurer: TextMeasurer
): FieldAwarePiece {
  const widthOf = (candidate: FieldAwarePiece): number =>
    widthOfPt(candidate, candidate.text, measurer);
  if (widthOf(piece) <= measurePt + 0.001) return piece;
  const labels = labelsOf(piece);
  const longest = labels.reduce((wide, label) => (label.length > wide.length ? label : wide));
  const graphemes = segmentGraphemes(piece.text.slice(longest.from, longest.from + longest.length));
  const cutAt = (kept: number): FieldAwarePiece => {
    const head = graphemes
      .slice(0, kept)
      .map((grapheme) => grapheme.text)
      .join('')
      .replace(/\s+$/u, '');
    return withLabelReplaced(piece, longest, `${head}${ELLIPSIS}`);
  };
  let low = 0;
  let high = graphemes.length - 1;
  while (low < high) {
    const middle = (low + high + 1) >> 1;
    if (widthOf(cutAt(middle)) <= measurePt + 0.001) low = middle;
    else high = middle - 1;
  }
  return cutAt(low);
}

interface LabelSlice {
  readonly index: number;
  readonly from: number;
  readonly length: number;
}

/** Each label a tag paints, by where it stands in the tag's text: the parts of a sequence, or one. */
function labelsOf(piece: FieldAwarePiece): readonly LabelSlice[] {
  const parts = piece.contentControlTag?.parts;
  if (parts === undefined) return [{ index: 0, from: 0, length: piece.text.length }];
  let from = 0;
  return parts.map((part, index) => {
    const slice = { index, from, length: part.length };
    from += part.length;
    return slice;
  });
}

function withLabelReplaced(
  piece: FieldAwarePiece,
  label: LabelSlice,
  text: string
): FieldAwarePiece {
  const tag = piece.contentControlTag!;
  const replaced = `${piece.text.slice(0, label.from)}${text}${piece.text.slice(label.from + label.length)}`;
  const parts = tag.parts?.map((part, index) =>
    index === label.index ? { ...part, length: text.length } : part
  );
  return {
    ...piece,
    text: replaced,
    contentControlTag: { ...tag, ...(parts === undefined ? {} : { parts }) },
  };
}
