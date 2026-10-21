// The state of an atom — a field, or a content control the reader meets as one unit — set on the
// atom's own painted span, in place.
//
// A control's chrome is a layer beside the text, not around it, so `data-hover` and `data-selected`
// on the chrome cannot reach the atom's text in a stylesheet. The atom carries them itself: the
// hover from the one layout answer the chrome's retint reads, the selection from the model's. The
// marks move as the pointer and the selection move, and nothing is repainted for them.

import type { EditorHover } from '../contracts/editor.ts';
import type { SemanticSelection } from '../layout/semantic-interaction.ts';

const HOVER_ATTRIBUTE = 'data-hover';
const SELECTED_ATTRIBUTE = 'data-selected';

/** What was last marked on a layer, so the next sync starts from bookkeeping, not queries. */
interface LayerMarks {
  key: string | null;
  marked: HTMLElement[];
}

const hoverByLayer = new WeakMap<HTMLElement, LayerMarks>();
const selectedByLayer = new WeakMap<HTMLElement, LayerMarks>();

/**
 * Mark the atom under the pointer: the atom control it is over, and the field it is on the glyphs
 * of. `domReplaced` says the spans were just rebuilt, so the marks must be resolved again.
 */
export function syncHoveredPieces(
  pagesLayer: HTMLElement,
  hover: EditorHover,
  options?: { readonly domReplaced?: boolean }
): void {
  const field = hover.field;
  const key =
    hover.controlId === null && field === null
      ? null
      : `${hover.controlId ?? ''}\u0000${field ? rangeKey(field.paragraphId, field.start, field.end) : ''}`;
  const marks = reset(hoverByLayer, pagesLayer, key, HOVER_ATTRIBUTE, options);
  if (!marks) return;
  for (const atom of pagesLayer.querySelectorAll<HTMLElement>('[data-piece-control]')) {
    if (atom.dataset.pieceControl === hover.controlId) mark(marks, atom, HOVER_ATTRIBUTE, '');
  }
  if (!field) return;
  for (const atom of pagesLayer.querySelectorAll<HTMLElement>('[data-field-atom]')) {
    if (spanCovers(atom, field.paragraphId, field.start, field.end)) {
      mark(marks, atom, HOVER_ATTRIBUTE, '');
    }
  }
}

/**
 * Mark the atom the selection is exactly: one field unit — the press on a field, or a range that
 * covers one and nothing else — or the whole content of an atom control. The selection highlight
 * is a tint under the text; a host drawing the atom as a chip needs to say which chip is the
 * subject of the toolbar and of its own inspector.
 */
export function syncSelectedPiece(
  pagesLayer: HTMLElement,
  selection: SemanticSelection,
  options?: { readonly domReplaced?: boolean }
): void {
  const range = singleParagraphRangeOf(selection);
  const key = range ? rangeKey(range.paragraphId, range.start, range.end) : null;
  const marks = reset(selectedByLayer, pagesLayer, key, SELECTED_ATTRIBUTE, options);
  if (!marks || !range) return;
  const candidates =
    range.end - range.start === 1
      ? '[data-field-atom], [data-piece-control]'
      : '[data-piece-control]';
  for (const atom of pagesLayer.querySelectorAll<HTMLElement>(candidates)) {
    if (spanCovers(atom, range.paragraphId, range.start, range.end)) {
      mark(marks, atom, SELECTED_ATTRIBUTE, 'true');
    }
  }
}

/**
 * Take down what the layer last marked, unless the same key meets the same DOM; the bookkeeping
 * to mark into, or null when there is nothing to do.
 */
function reset(
  byLayer: WeakMap<HTMLElement, LayerMarks>,
  pagesLayer: HTMLElement,
  key: string | null,
  attribute: string,
  options: { readonly domReplaced?: boolean } | undefined
): LayerMarks | null {
  let marks = byLayer.get(pagesLayer);
  if (!marks) {
    marks = { key: null, marked: [] };
    byLayer.set(pagesLayer, marks);
  }
  if (!options?.domReplaced && key === marks.key) return null;
  for (const marked of marks.marked) marked.removeAttribute(attribute);
  marks.marked = [];
  marks.key = key;
  return key === null ? null : marks;
}

function mark(marks: LayerMarks, atom: HTMLElement, attribute: string, value: string): void {
  atom.setAttribute(attribute, value);
  marks.marked.push(atom);
}

/**
 * Whether a painted span publishes exactly this range. The paragraph id is compared in JS, never
 * written into a selector: it is built from a part name, which comes out of the file.
 */
function spanCovers(span: HTMLElement, paragraphId: string, start: number, end: number): boolean {
  return (
    span.dataset.paragraphId === paragraphId &&
    Number(span.dataset.start) === start &&
    Number(span.dataset.end) === end
  );
}

/** NUL-joined: a paragraph id cannot carry one, so the key cannot collide across ids. */
function rangeKey(paragraphId: string, start: number, end: number): string {
  return `${paragraphId}\u0000${start}\u0000${end}`;
}

/** A non-empty selection inside one paragraph, in order, or null. */
function singleParagraphRangeOf(
  selection: SemanticSelection
): { readonly paragraphId: string; readonly start: number; readonly end: number } | null {
  const { anchor, head } = selection;
  if (anchor.paragraphId !== head.paragraphId || anchor.offset === head.offset) return null;
  return {
    paragraphId: head.paragraphId,
    start: Math.min(anchor.offset, head.offset),
    end: Math.max(anchor.offset, head.offset),
  };
}
