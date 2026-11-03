// What a selection takes whole, from the boundary geometry layout publishes: a UNIT — a content
// control showing its prompt, or one the host made an atom — and a frame the host takes whole.
// `surface-pointer.ts` owns the gesture.

import { contentControlAtPoint, type SemanticHit } from '../layout/semantic-hit-test.ts';
import {
  documentOrder,
  paragraphTextFromLayout,
  type SemanticPosition,
  type SemanticSelection,
} from '../layout/semantic-interaction.ts';
import type {
  BlockFragmentRecord,
  ContentControlBoundaryRecord,
  LayoutBox,
  ParagraphFragmentRecord,
  SemanticLayout,
} from '../layout/semantic-records.ts';
import { MAX_TABLE_NESTING } from '../layout/semantic-table.ts';

/** A model range a press selects, in document order. */
export interface PositionRange {
  readonly from: SemanticPosition;
  readonly to: SemanticPosition;
}

function overlapArea(a: LayoutBox, b: LayoutBox): number {
  const x0 = Math.max(a.x, b.x);
  const y0 = Math.max(a.y, b.y);
  const x1 = Math.min(a.x + a.width, b.x + b.width);
  const y1 = Math.min(a.y + a.height, b.y + b.height);
  if (x1 <= x0 || y1 <= y0) return 0;
  return (x1 - x0) * (y1 - y0);
}

/** True when boxes share area, or a zero-width control box sits on `inner`. */
function boxesTouch(outer: LayoutBox, inner: LayoutBox): boolean {
  if (overlapArea(outer, inner) > 0) return true;
  if (outer.width === 0) {
    return (
      inner.x <= outer.x &&
      outer.x <= inner.x + inner.width &&
      !(outer.y + outer.height < inner.y || inner.y + inner.height < outer.y)
    );
  }
  return false;
}

function comparePositions(
  layout: SemanticLayout,
  a: SemanticPosition,
  b: SemanticPosition
): number {
  const order = documentOrder(layout);
  const rankA = order.indexOf(a.paragraphId);
  const rankB = order.indexOf(b.paragraphId);
  if (rankA !== rankB) return rankA - rankB;
  return a.offset - b.offset;
}

/**
 * Document-order paragraph walk that descends table rows/cells within {@link MAX_TABLE_NESTING}.
 *
 * Header-repeat rows are skipped so each caret stop is considered once. Nested tables inside
 * cells share the same bound layout already uses when refusing deeper nesting.
 */
function visitParagraphBlocks(
  blocks: readonly BlockFragmentRecord[],
  depth: number,
  visit: (paragraph: ParagraphFragmentRecord) => void
): void {
  if (depth > MAX_TABLE_NESTING) return;
  for (const block of blocks) {
    if (block.kind === 'paragraph') {
      visit(block);
      continue;
    }
    if (depth >= MAX_TABLE_NESTING) continue;
    for (const row of block.rows) {
      if (row.isHeaderRepeat) continue;
      for (const cell of row.cells) visitParagraphBlocks(cell.blocks, depth + 1, visit);
    }
  }
}

/** Whether the reader meets a control as one unit: it shows its prompt, or the host made it an atom. */
export function isUnitControl(control: ContentControlBoundaryRecord): boolean {
  return control.placeholder || control.atom;
}

/**
 * Full addressable range of a unit control from layout-published boundary geometry.
 *
 * Block / row / cell controls cover whole paragraphs whose boxes touch a fragment (including
 * paragraphs nested under tables and nested block-control content that flattened into the
 * flow). Inline controls take the UTF-16 span of every painted span that touches a fragment.
 */
export function unitSelectionRange(
  layout: SemanticLayout,
  control: ContentControlBoundaryRecord
): PositionRange | null {
  return isUnitControl(control) ? controlRangeOf(layout, control) : null;
}

/** A control's full addressable range, by the same geometry as {@link unitSelectionRange}. */
function controlRangeOf(
  layout: SemanticLayout,
  control: ContentControlBoundaryRecord
): PositionRange | null {
  if (control.fragments.length === 0) return null;

  let from: SemanticPosition | null = null;
  let to: SemanticPosition | null = null;

  const consider = (start: SemanticPosition, end: SemanticPosition): void => {
    if (!from || comparePositions(layout, start, from) < 0) from = start;
    if (!to || comparePositions(layout, end, to) > 0) to = end;
  };

  // Non-inline controls select whole paragraphs — row/cell wrappers are structural units
  // the same way block wrappers are, so a mid-cell partial range must still absorb.
  const wholeParagraphs = control.level !== 'inline';

  for (const fragment of control.fragments) {
    const page = layout.pages[fragment.pageIndex];
    if (!page) continue;
    visitParagraphBlocks(page.fragments, 0, (block) => {
      if (!boxesTouch(fragment.box, block.box)) return;

      if (wholeParagraphs) {
        const length = paragraphTextFromLayout(layout, block.paragraphId).length;
        consider(
          { paragraphId: block.paragraphId, offset: 0 },
          { paragraphId: block.paragraphId, offset: length }
        );
        return;
      }

      for (const line of block.lines) {
        for (const span of line.spans) {
          if (!boxesTouch(fragment.box, span.box)) continue;
          consider(
            { paragraphId: span.range.paragraphId, offset: span.range.start },
            { paragraphId: span.range.paragraphId, offset: span.range.end }
          );
        }
      }
    });
  }

  if (!from || !to) return null;
  return { from, to };
}

function controlById(
  layout: SemanticLayout,
  controlId: string
): ContentControlBoundaryRecord | null {
  return (layout.contentControls ?? []).find((control) => control.id === controlId) ?? null;
}

export function unitAtHit(
  layout: SemanticLayout,
  hit: SemanticHit
): ContentControlBoundaryRecord | null {
  const id = hit.contentControlId;
  if (!id) {
    // Hit records may omit the id when resolved through a scoped story path — fall back to
    // geometry so unit expansion still works inside headers/notes.
    const point = { x: hit.caret.x, y: hit.caret.y + hit.caret.height / 2 };
    const control = contentControlAtPoint(layout, hit.pageIndex, point);
    return control && isUnitControl(control) ? control : null;
  }
  const control = controlById(layout, id);
  return control && isUnitControl(control) ? control : null;
}

/**
 * Expand a selection so every unit control it touches is selected whole — a range over any unit,
 * and a caret inside a prompt — and so is every frame taken whole whose edge it crosses.
 *
 * Used by pointer drag / shift-click, and exported so keyboard extend paths can share the
 * same atomic rule without duplicating geometry walks.
 */
export function absorbUnitControls(
  layout: SemanticLayout,
  selection: SemanticSelection
): SemanticSelection {
  const controls = (layout.contentControls ?? []).filter(
    (control) => isUnitControl(control) || control.whole
  );
  if (controls.length === 0) return selection;

  const headIsEnd = comparePositions(layout, selection.anchor, selection.head) <= 0;
  let taken: PositionRange = headIsEnd
    ? { from: selection.anchor, to: selection.head }
    : { from: selection.head, to: selection.anchor };

  let changed = true;
  while (changed) {
    changed = false;
    for (const control of controls) {
      const range = rangeTakenIn(layout, control, taken);
      if (!range) continue;
      const grown = spanOf(layout, taken, range);
      changed ||= !sameRange(layout, grown, taken);
      taken = grown;
    }
  }

  return headIsEnd
    ? { anchor: taken.from, head: taken.to }
    : { anchor: taken.to, head: taken.from };
}

/** The range of a control a selection takes in whole, or null where it leaves the control alone. */
function rangeTakenIn(
  layout: SemanticLayout,
  control: ContentControlBoundaryRecord,
  selection: PositionRange
): PositionRange | null {
  const range = controlRangeOf(layout, control);
  if (!range) return null;
  const takesIn = isUnitControl(control)
    ? touchesUnit(layout, control, range, selection)
    : reachesFromOutside(layout, range, selection);
  return takesIn ? range : null;
}

/** A range over a unit touches it anywhere; a caret only a prompt, which it stands inside. */
function touchesUnit(
  layout: SemanticLayout,
  control: ContentControlBoundaryRecord,
  range: PositionRange,
  selection: PositionRange
): boolean {
  if (comparePositions(layout, selection.from, selection.to) !== 0) {
    return overlaps(layout, range, selection);
  }
  return (
    control.placeholder &&
    comparePositions(layout, selection.from, range.from) >= 0 &&
    comparePositions(layout, selection.from, range.to) < 0
  );
}

/** A frame taken whole is taken in by a selection that overlaps it and does not stay inside it. */
function reachesFromOutside(
  layout: SemanticLayout,
  range: PositionRange,
  selection: PositionRange
): boolean {
  const inside =
    comparePositions(layout, range.from, selection.from) <= 0 &&
    comparePositions(layout, selection.to, range.to) <= 0;
  return overlaps(layout, range, selection) && !inside;
}

function overlaps(layout: SemanticLayout, a: PositionRange, b: PositionRange): boolean {
  return comparePositions(layout, b.from, a.to) < 0 && comparePositions(layout, b.to, a.from) > 0;
}

function spanOf(layout: SemanticLayout, a: PositionRange, b: PositionRange): PositionRange {
  return {
    from: comparePositions(layout, b.from, a.from) < 0 ? b.from : a.from,
    to: comparePositions(layout, b.to, a.to) > 0 ? b.to : a.to,
  };
}

function sameRange(layout: SemanticLayout, a: PositionRange, b: PositionRange): boolean {
  return (
    comparePositions(layout, a.from, b.from) === 0 && comparePositions(layout, a.to, b.to) === 0
  );
}

/**
 * Where a caret move that lands INSIDE an atom goes: to the atom's edge on the side it was moving
 * to, so the caret steps over it in one step, as it does over a field. A prompt is absorbed
 * instead ({@link absorbUnitControls}): its text is replaced by the first key.
 */
export function stepOverAtoms(
  layout: SemanticLayout,
  position: SemanticPosition,
  backward: boolean
): SemanticPosition {
  for (const control of layout.contentControls ?? []) {
    if (!control.atom) continue;
    const range = unitSelectionRange(layout, control);
    if (!range) continue;
    const inside =
      comparePositions(layout, position, range.from) > 0 &&
      comparePositions(layout, position, range.to) < 0;
    if (inside) return backward ? range.from : range.to;
  }
  return position;
}

const atomIdsByLayout = new WeakMap<SemanticLayout, ReadonlySet<string>>();
const wholeIdsByLayout = new WeakMap<SemanticLayout, ReadonlySet<string>>();

/** The controls the host made atoms, once per layout. */
export function atomControlIdsOf(layout: SemanticLayout): ReadonlySet<string> {
  return controlIdsOnce(atomIdsByLayout, layout, (control) => control.atom);
}

/** The frames the host made pieces taken whole, once per layout. */
export function wholeControlIdsOf(layout: SemanticLayout): ReadonlySet<string> {
  return controlIdsOnce(wholeIdsByLayout, layout, (control) => control.whole);
}

function controlIdsOnce(
  byLayout: WeakMap<SemanticLayout, ReadonlySet<string>>,
  layout: SemanticLayout,
  keeps: (control: ContentControlBoundaryRecord) => boolean
): ReadonlySet<string> {
  let ids = byLayout.get(layout);
  if (!ids) {
    ids = new Set((layout.contentControls ?? []).filter(keeps).map((each) => each.id));
    byLayout.set(layout, ids);
  }
  return ids;
}

/**
 * The range a press selects an atom by, or null for a control that is no atom. The keys hold it
 * whole by the control, not by the range: an empty atom's whole is no range at all, and it is still
 * the one the next key removes.
 */
export function atomRangeOf(layout: SemanticLayout, controlId: string): PositionRange | null {
  if (!atomControlIdsOf(layout).has(controlId)) return null;
  const control = controlById(layout, controlId);
  return control ? unitSelectionRange(layout, control) : null;
}
