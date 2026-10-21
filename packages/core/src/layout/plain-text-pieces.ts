// The content of a control a view lays out as ONE piece — its prompt, or the host's atom — when
// that content is contiguous plain text: what the walk emitted for it, from one index on.

import type { FieldAwarePiece } from './field-pieces.ts';

/** What a piece of plain text may carry; anything more is something the file marks. */
const PLAIN_PIECE_KEYS: ReadonlySet<string> = new Set(['text', 'props', 'style', 'start', 'end']);

/** Contiguous plain text, read from `from` to the end of `pieces`. */
export interface PlainText {
  readonly first: FieldAwarePiece;
  readonly text: string;
  readonly end: number;
}

/**
 * The text of `pieces` from `from` on, or null when there is none or any piece is not plain text:
 * tracked, linked, a field, a drawing, a break, or another control's tag.
 */
export function plainTextFrom(pieces: readonly FieldAwarePiece[], from: number): PlainText | null {
  const first = pieces[from];
  if (!first) return null;
  let text = '';
  let end = first.start;
  for (let index = from; index < pieces.length; index += 1) {
    const piece = pieces[index]!;
    if (piece.start !== end || piece.end - piece.start !== piece.text.length) return null;
    if (!Object.keys(piece).every((key) => PLAIN_PIECE_KEYS.has(key))) return null;
    text += piece.text;
    end = piece.end;
  }
  return { first, text, end };
}
