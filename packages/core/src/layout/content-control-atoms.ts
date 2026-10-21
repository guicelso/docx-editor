// The content control a host makes an atom: the reader meets it as one unit, as a control showing
// its placeholder. Its content keeps its model range; when that content is plain text in one face,
// the walk lays it out as ONE projected piece with the atom's room and mark, and its own text — so
// the text layout reads back, the clipboard and every offset still agree with the store.

import type { OoxmlElement } from '@docx-editor.dev/core/store';
import type {
  ContentControlPieceDisplay,
  PieceAtom,
} from '../contracts/editor-content-control-view.ts';
import { contentControlPropertiesOf as storedPropertiesOf } from '../store/package/content-control-nodes.ts';
import type { BlockControlEdges } from '../store/store/block-control-edges.ts';
import { contentControlSubjectOf } from './content-control-properties.ts';
import type { FieldAwarePiece } from './field-pieces.ts';
import { pieceAtomRecordOf } from './piece-names.ts';
import { plainTextFrom } from './plain-text-pieces.ts';
import { runStylesEqual } from './run-style.ts';

/** The host's atom for a control, or null. A control showing its placeholder shows its prompt. */
export function contentControlAtomOf(
  control: OoxmlElement,
  display: ContentControlPieceDisplay | undefined
): PieceAtom | null {
  if (!display || storedPropertiesOf(control).showingPlaceholder) return null;
  const piece = display.pieceOf(contentControlSubjectOf(control));
  return piece !== null && 'atom' in piece ? piece.atom : null;
}

/** The atom whose only paragraph this is, innermost first, as the prompt of a block control. */
export function blockContentControlAtomOf(
  edges: BlockControlEdges,
  display: ContentControlPieceDisplay | undefined
): Readonly<{ controlId: string; atom: PieceAtom }> | null {
  if (!display) return null;
  for (let index = edges.opens.length - 1; index >= 0; index -= 1) {
    const control = edges.opens[index]!;
    if (!edges.closes.includes(control)) continue;
    const atom = contentControlAtomOf(control, display);
    if (atom !== null) return { controlId: control.id, atom };
  }
  return null;
}

/**
 * Lay the atom's content, `pieces` from `from` on, out as one piece. Text in more than one face
 * keeps its pieces — one span paints one face — and so does anything the file marks.
 */
export function projectContentControlAtom(
  pieces: FieldAwarePiece[],
  from: number,
  controlId: string,
  atom: PieceAtom
): void {
  const plain = plainTextFrom(pieces, from);
  if (!plain || !inOneFace(pieces.slice(from))) return;
  pieces.splice(from, pieces.length - from, {
    text: plain.text,
    props: plain.first.props,
    style: plain.first.style,
    start: plain.first.start,
    end: plain.end,
    projected: true,
    pieceAtom: pieceAtomRecordOf(atom, controlId),
  });
}

function inOneFace(pieces: readonly FieldAwarePiece[]): boolean {
  return pieces.every((piece) => runStylesEqual(piece.style, pieces[0]!.style));
}
