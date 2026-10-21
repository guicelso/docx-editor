// Pieces the page paints in the room layout reserved for them: a content-control tag, the parts of a
// sequence of labels, an atom, and the slot of a mark. Split out of `semantic-paint.ts` so a piece's
// geometry lives next to the one function that measures it (`piece-names.ts`).

import { contentControlTagRoomOf } from '../layout/content-control-tags.ts';
import {
  NO_PIECE_ROOM,
  PIECE_MARK_GLYPH_EM,
  PIECE_MARK_SPACE_EM,
  pieceRoomOf,
  type PieceNameRecord,
  type PieceRoom,
} from '../layout/piece-names.ts';
import type { StyleSpanRecord } from '@docx-editor.dev/core/layout';

/**
 * Fills one painted mark slot (`data-piece-mark`), sized by the engine, with what the host draws
 * for `mark`. Called for every slot each time a page is painted. @public
 */
export type MarkPainter = (mark: string, slot: HTMLElement) => void;

/** The host's names on a painted element, for its stylesheet; nothing for a name it did not give. */
export function paintPieceName(element: HTMLElement, name: PieceNameRecord | undefined): void {
  if (name?.tone !== undefined) element.dataset.pieceTone = name.tone;
  if (name?.form !== undefined) element.dataset.pieceForm = name.form;
  if (name?.state !== undefined) element.dataset.pieceState = name.state;
  if (name?.role !== undefined) element.dataset.pieceRole = name.role;
}

/**
 * A piece keeps EXACTLY the advance layout reserved — the caret and every following glyph are
 * placed from that number. The gap between neighbouring pieces is a margin on each side, outside
 * the fill: a transparent border there clipped the fill's corners square. The mark's slot is in the
 * content box, before the text. Its height is the run's band, like any span's.
 */
function paintPieceRoom(element: HTMLElement, width: number, room: PieceRoom, scale: number): void {
  const css = element.style;
  css.width = `${(width - 2 * room.gapPt) * scale}px`;
  css.marginLeft = `${room.gapPt * scale}px`;
  css.marginRight = `${room.gapPt * scale}px`;
  css.paddingLeft = `${room.padPt * scale}px`;
  css.paddingRight = `${room.padPt * scale}px`;
  css.textAlign = 'center';
  // No `overflow: hidden`: on an inline-block it moves the baseline to the bottom edge
  // (CSS 2.1 §10.8.1) and lifts the piece off the line. The text fits by construction —
  // layout measured this very text for the advance.
  css.whiteSpace = 'nowrap';
}

/**
 * A view-only content-control tag drawn as its own chip. Colour, fill and outline belong to the
 * stylesheet: `.docx-cc-tag` draws a neutral chip from the `--doc-*` tokens,
 * `data-cc-tag-variant="text"` draws no fill, and a host restyles it by the names it chose. The
 * data attributes are also the host's click target. A tag of an atom names its control, as the
 * atom does, so the unit's state reaches it in place.
 */
export function paintContentControlTagChip(
  element: HTMLElement,
  span: StyleSpanRecord,
  scale: number
): void {
  const tag = span.contentControlTag!;
  element.classList.add('docx-cc-tag');
  paintPieceRoom(element, span.box.width, contentControlTagRoomOf(span.style, tag), scale);
  element.dataset.ccTagControl = tag.controlId;
  element.dataset.ccTagEdge = tag.edge;
  if (tag.atom) element.dataset.pieceControl = tag.controlId;
  paintPieceName(element, tag.name);
  if (tag.variant !== undefined) element.dataset.ccTagVariant = tag.variant;
}

/**
 * A sequence paints each of its labels as a part of its own, with the room that label keeps and
 * the names, mark and variant it chose, so the stylesheet colours each part as it would a tag
 * alone. The parts share the tag's band, and together fill the advance layout measured.
 */
export function mountContentControlTagParts(
  document: Document,
  host: HTMLElement,
  span: StyleSpanRecord,
  scale: number,
  painter: MarkPainter | undefined
): void {
  let from = 0;
  for (const part of span.contentControlTag!.parts!) {
    const element = document.createElement('span');
    const room = contentControlTagRoomOf(span.style, part);
    element.classList.add('docx-cc-tag');
    element.style.display = 'inline-block';
    element.style.height = '100%';
    element.style.marginLeft = `${room.gapPt * scale}px`;
    element.style.marginRight = `${room.gapPt * scale}px`;
    element.style.paddingLeft = `${room.padPt * scale}px`;
    element.style.paddingRight = `${room.padPt * scale}px`;
    paintPieceName(element, part.name);
    if (part.variant !== undefined) element.dataset.ccTagVariant = part.variant;
    element.textContent = span.text.slice(from, from + part.length); // SAFE: textContent, never innerHTML
    mountPieceMark(document, element, part.mark, span, scale, painter);
    from += part.length;
    host.append(element);
  }
}

/**
 * An atom: a field with the host's face, or a content control met as one unit. It is painted in
 * the chip's room, with the host's names, and an atom control names its control so its state can
 * be set on it in place.
 */
export function paintPieceAtom(element: HTMLElement, span: StyleSpanRecord, scale: number): void {
  const atom = span.pieceAtom!;
  paintPieceRoom(element, span.box.width, pieceRoomOf(span.style, atom.mark), scale);
  paintPieceName(element, atom.name);
  if (atom.controlId !== undefined) element.dataset.pieceControl = atom.controlId;
}

/** The mark a span draws before its text: an atom's, or a chip's. A sequence's parts draw their own. */
export function markOfSpan(span: StyleSpanRecord): string | undefined {
  if (span.pieceAtom) return span.pieceAtom.mark;
  return span.contentControlTag?.parts ? undefined : span.contentControlTag?.mark;
}

/**
 * The slot of a mark, as wide as layout measured it, first in `host`, handed to the host's painter.
 * Decorative: the piece's name is beside it. Without a painter it stays empty.
 */
export function mountPieceMark(
  document: Document,
  host: HTMLElement,
  mark: string | undefined,
  span: StyleSpanRecord,
  scale: number,
  painter: MarkPainter | undefined
): void {
  if (mark === undefined) return;
  const em = span.style.fontSizePt * scale;
  const slot = document.createElement('span');
  slot.className = 'docx-piece-mark';
  slot.dataset.pieceMark = mark;
  slot.dataset.docxMarker = '';
  slot.setAttribute('aria-hidden', 'true');
  slot.setAttribute('contenteditable', 'false');
  slot.style.width = `${PIECE_MARK_GLYPH_EM * em}px`;
  slot.style.height = `${PIECE_MARK_GLYPH_EM * em}px`;
  slot.style.marginRight = `${PIECE_MARK_SPACE_EM * em}px`;
  host.prepend(slot);
  painter?.(mark, slot);
}

/** The margin a piece keeps on its left, inside its advance; none for any other span. */
export function pieceGapPt(span: StyleSpanRecord): number {
  if (span.contentControlTag)
    return contentControlTagRoomOf(span.style, span.contentControlTag).gapPt;
  return span.pieceAtom ? pieceRoomOf(span.style, span.pieceAtom.mark).gapPt : NO_PIECE_ROOM.gapPt;
}
