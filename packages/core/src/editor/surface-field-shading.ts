// Word's "field shading: when selected", which is a CARET question, not a layout one.
//
// The mode shades a field only while the insertion point is inside it. Resolving that in layout
// would put the caret into the per-block cache key and remeasure the document on every arrow
// press; resolving it in paint would rebuild spans just as often. Both are the wrong lane for a
// background colour.
//
// So layout marks which spans ARE fields (`data-field-atom`, painted once) and this toggles one
// class as the caret moves — the same division the open review item already uses. Cost is at
// most one query per caret move, against a document that never relayouts.

const ACTIVE_CLASS = 'docx-field-atom--active';
/** What paint adds only where shading is enabled for that field. */
const SHADABLE_CLASS = 'docx-field-atom';

/** Where the caret is, in the model's own addressing. */
export interface FieldShadingCaret {
  readonly paragraphId: string;
  readonly offset: number;
}

/**
 * What this module last did to a layer, so the next sync starts from bookkeeping, not queries.
 *
 * `marked` is every element carrying {@link ACTIVE_CLASS} — this function is the class's only
 * writer, so the list is complete and the removal sweep needs no DOM query. `caretKey` is the
 * caret the marks were resolved for: a sync that lands on the same caret against the same DOM
 * has nothing to do, which is the arrow-key-then-mirror double call every caret move used to
 * pay twice. Keyed weakly so a detached surface holds nothing.
 */
interface LayerShadingState {
  caretKey: string | null;
  marked: HTMLElement[];
}
const shadingStateByLayer = new WeakMap<HTMLElement, LayerShadingState>();

/**
 * Move the "caret is in this field" mark to whichever field atom holds `caret`.
 *
 * `caret` is null when the selection is not collapsed, and every mark then comes off: a range
 * draws its own highlight, and a second background under one end of it reads as a second
 * selection. Focus and IME composition are deliberately NOT part of that test — Word keeps a
 * field shaded while the caret is in it, and losing the shading on every blur would flicker it
 * away each time the user reached for the toolbar.
 *
 * `domReplaced` says the spans the marks live on were just rebuilt — a paint — so the caret
 * being where it already was proves nothing and the marks must be re-resolved. Every other
 * caller leaves it off and an unmoved caret costs no query at all.
 */
export function syncActiveFieldShading(
  pagesLayer: HTMLElement,
  caret: FieldShadingCaret | null,
  options?: { readonly domReplaced?: boolean }
): void {
  let state = shadingStateByLayer.get(pagesLayer);
  if (!state) {
    state = { caretKey: null, marked: [] };
    shadingStateByLayer.set(pagesLayer, state);
  }
  // NUL-joined: a paragraph id cannot carry one, so the key cannot collide across ids.
  const caretKey = caret ? `${caret.paragraphId}\u0000${caret.offset}` : null;
  if (!options?.domReplaced && caretKey === state.caretKey) return;
  // Removing a class from a node paint already replaced is harmless; a RETAINED page keeps
  // its nodes, and those are exactly the marks that must come off here.
  for (const marked of state.marked) marked.classList.remove(ACTIVE_CLASS);
  state.marked = [];
  state.caretKey = caretKey;
  if (!caret) return;

  // `.docx-field-atom`, NOT `[data-field-atom]`. The attribute marks every field result so a
  // host can find them; the CLASS is what paint adds only where shading is actually enabled.
  // Keying off the attribute meant the caret shaded a field that `fieldShading: 'never'` — or
  // the document's own `w:doNotShadeFormData` — had just said never to shade, because the
  // stylesheet paints `--active` unconditionally. The caret must not be able to overrule the
  // decision paint already made.
  //
  // It is also the cheaper selector. This runs on the keystroke path, and a class match is
  // what browsers optimise; an attribute match over every materialized page is not.
  //
  // The paragraph id is compared in JS rather than written into the selector. It is
  // engine-minted, but it is built from a PART NAME, and part names come out of the file — so
  // it is attacker-influenced text, and the only way to interpolate it safely into a selector
  // is not to.
  const candidates = pagesLayer.querySelectorAll<HTMLElement>(`.${SHADABLE_CLASS}`);
  for (const candidate of candidates) {
    if (candidate.dataset.paragraphId !== caret.paragraphId) continue;
    const start = Number(candidate.dataset.start);
    const end = Number(candidate.dataset.end);
    if (!Number.isFinite(start) || !Number.isFinite(end)) continue;
    // Half-open on the left, INCLUSIVE on the right: a field is one model unit, so a caret
    // resting at either edge is a caret Word considers inside it. Excluding the trailing edge
    // made the shading flicker off as the caret arrived at the field from the left.
    if (caret.offset < start || caret.offset > end) continue;
    // Every span, not the first. Line breaking splits a field's result at its spaces like any
    // other text, and all of them publish the same model range — marking one shaded half a
    // cross-reference while `always` (resolved in paint, per span) shaded all of it.
    candidate.classList.add(ACTIVE_CLASS);
    state.marked.push(candidate);
  }
}

/** A selection one model unit wide, which is what a selected field atom is. */
function oneUnitRangeOf(selection: {
  readonly anchor: FieldShadingCaret;
  readonly head: FieldShadingCaret;
}): { readonly paragraphId: string; readonly start: number; readonly end: number } | null {
  const { anchor, head } = selection;
  if (anchor.paragraphId !== head.paragraphId || Math.abs(head.offset - anchor.offset) !== 1) {
    return null;
  }
  const start = Math.min(anchor.offset, head.offset);
  return { paragraphId: head.paragraphId, start, end: start + 1 };
}

const SELECTED_ATTRIBUTE = 'data-selected';
const selectedByLayer = new WeakMap<HTMLElement, LayerShadingState>();

/**
 * Mark the field atom the selection is exactly — the press on a field, or a range that covers
 * one and nothing else — as `data-selected`, the attribute a selected content control carries.
 *
 * A one-unit range is painted by the selection highlight over a field's whole result, but the
 * highlight is a tint under the text: a host styling the field as a chip needs to say which chip
 * is the subject of the toolbar and of its own inspector. Same bookkeeping as the caret's mark.
 */
export function syncSelectedField(
  pagesLayer: HTMLElement,
  selection: {
    readonly anchor: FieldShadingCaret;
    readonly head: FieldShadingCaret;
  },
  options?: { readonly domReplaced?: boolean }
): void {
  const range = oneUnitRangeOf(selection);
  let state = selectedByLayer.get(pagesLayer);
  if (!state) {
    state = { caretKey: null, marked: [] };
    selectedByLayer.set(pagesLayer, state);
  }
  const key = range ? `${range.paragraphId}\u0000${range.start}\u0000${range.end}` : null;
  if (!options?.domReplaced && key === state.caretKey) return;
  for (const marked of state.marked) marked.removeAttribute(SELECTED_ATTRIBUTE);
  state.marked = [];
  state.caretKey = key;
  if (!range) return;
  for (const candidate of pagesLayer.querySelectorAll<HTMLElement>('[data-field-atom]')) {
    if (candidate.dataset.paragraphId !== range.paragraphId) continue;
    if (Number(candidate.dataset.start) !== range.start) continue;
    if (Number(candidate.dataset.end) !== range.end) continue;
    candidate.setAttribute(SELECTED_ATTRIBUTE, 'true');
    state.marked.push(candidate);
  }
}
