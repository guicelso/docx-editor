// The host's names for a piece, checked once where they enter layout, and the room a piece keeps
// around its text: the one geometry a chip and an atom share, which layout adds to the measured
// text and paint draws exactly.

import type { PieceAtom, PieceName } from '../contracts/editor-content-control-view.ts';
import type { ResolvedRunStyle } from './run-style.ts';

const NAME = /^[A-Za-z][A-Za-z0-9-]{0,31}$/;

/** A host's names that passed the pattern: they land in attributes. @public */
export interface PieceNameRecord {
  /** `PieceName.tone`, checked. */
  readonly tone?: string;
  /** `PieceName.form`, checked. */
  readonly form?: string;
  /** `PieceName.state`, checked. */
  readonly state?: string;
  /** `PieceName.role`, checked. */
  readonly role?: string;
}

/** A piece laid out as one unit, as layout carries it to paint and to the hit test. @public */
export interface PieceAtomRecord {
  readonly name: PieceNameRecord;
  /** The mark measured before the text, already checked. */
  readonly mark?: string;
  /** The content control this atom stands for; absent for a field. */
  readonly controlId?: string;
}

/** One name the host chose, or undefined when it may not land in an attribute. */
export function checkedPieceNameOf(value: string | undefined): string | undefined {
  return value !== undefined && NAME.test(value) ? value : undefined;
}

/** The names a host chose that may land in attributes; none when it chose none that may. */
export function pieceNameRecordOf(name: PieceName | undefined): PieceNameRecord | undefined {
  const tone = checkedPieceNameOf(name?.tone);
  const form = checkedPieceNameOf(name?.form);
  const state = checkedPieceNameOf(name?.state);
  const role = checkedPieceNameOf(name?.role);
  if (tone === undefined && form === undefined && state === undefined && role === undefined) {
    return undefined;
  }
  return {
    ...(tone === undefined ? {} : { tone }),
    ...(form === undefined ? {} : { form }),
    ...(state === undefined ? {} : { state }),
    ...(role === undefined ? {} : { role }),
  };
}

/** The atom a host's answer makes, with its names and its mark checked. */
export function pieceAtomRecordOf(atom: PieceAtom, controlId?: string): PieceAtomRecord {
  const mark = checkedPieceNameOf(atom.mark);
  return {
    name: pieceNameRecordOf(atom) ?? {},
    ...(mark === undefined ? {} : { mark }),
    ...(controlId === undefined ? {} : { controlId }),
  };
}

/** The transparent gap that keeps two neighbouring chips apart, inside each chip's advance. */
const PIECE_GAP_PT = 1.2;
/** The padding between the piece's fill and its text, in ems of the piece's own size. */
const PIECE_PAD_EM = 0.35;
/** The mark's glyph box, and the space between it and the text, in ems of the text beside it. */
export const PIECE_MARK_GLYPH_EM = 0.85;
export const PIECE_MARK_SPACE_EM = 0.3;

/**
 * The room a piece keeps around its text, in points: the gap and the padding on EACH side, and the
 * mark's slot before the text. Geometry, not characters — layout adds it to the measured text and
 * paint draws exactly it, so the text sits in the space reserved for it whatever font paints it.
 */
export interface PieceRoom {
  readonly gapPt: number;
  readonly padPt: number;
  readonly markPt: number;
}

export const NO_PIECE_ROOM: PieceRoom = Object.freeze({ gapPt: 0, padPt: 0, markPt: 0 });

/** The room of a chip or an atom in `style`, with a mark slot when it has a mark. */
export function pieceRoomOf(
  style: Pick<ResolvedRunStyle, 'fontSizePt'>,
  mark: string | undefined
): PieceRoom {
  return {
    gapPt: PIECE_GAP_PT,
    padPt: style.fontSizePt * PIECE_PAD_EM,
    markPt: mark === undefined ? 0 : style.fontSizePt * (PIECE_MARK_GLYPH_EM + PIECE_MARK_SPACE_EM),
  };
}

/** What a room adds to its text's measured width. */
export function pieceRoomWidthPt(room: PieceRoom): number {
  return 2 * (room.gapPt + room.padPt) + room.markPt;
}
