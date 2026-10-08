// What one press selects as a UNIT: a content control showing its prompt, from the boundary
// geometry layout publishes. Split out of `surface-pointer.ts`, which owns the gesture.

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

/**
 * Full addressable range of a placeholder control from layout-published boundary geometry.
 *
 * Block / row / cell controls cover whole paragraphs whose boxes touch a fragment (including
 * paragraphs nested under tables and nested block-control content that flattened into the
 * flow). Inline controls take the UTF-16 span of every painted span that touches a fragment.
 */
export function placeholderSelectionRange(
  layout: SemanticLayout,
  control: ContentControlBoundaryRecord
): PositionRange | null {
  if (!control.placeholder || control.fragments.length === 0) return null;

  let from: SemanticPosition | null = null;
  let to: SemanticPosition | null = null;

  const consider = (start: SemanticPosition, end: SemanticPosition): void => {
    if (!from || comparePositions(layout, start, from) < 0) from = start;
    if (!to || comparePositions(layout, end, to) > 0) to = end;
  };

  // Non-inline placeholders select whole paragraphs — row/cell wrappers are structural units
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

export function placeholderAtHit(
  layout: SemanticLayout,
  hit: SemanticHit
): ContentControlBoundaryRecord | null {
  const id = hit.contentControlId;
  if (!id) {
    // Hit records may omit the id when resolved through a scoped story path — fall back to
    // geometry so placeholder expansion still works inside headers/notes.
    const point = { x: hit.caret.x, y: hit.caret.y + hit.caret.height / 2 };
    const control = contentControlAtPoint(layout, hit.pageIndex, point);
    return control?.placeholder ? control : null;
  }
  const control = controlById(layout, id);
  return control?.placeholder ? control : null;
}

/**
 * Expand a selection so every `w:showingPlcHdr` control it touches is selected as a unit.
 *
 * Used by pointer drag / shift-click, and exported so keyboard extend paths can share the
 * same atomic rule without duplicating geometry walks.
 */
export function absorbPlaceholderControls(
  layout: SemanticLayout,
  selection: SemanticSelection
): SemanticSelection {
  const controls = layout.contentControls ?? [];
  if (controls.length === 0) return selection;

  let from =
    comparePositions(layout, selection.anchor, selection.head) <= 0
      ? selection.anchor
      : selection.head;
  let to =
    comparePositions(layout, selection.anchor, selection.head) <= 0
      ? selection.head
      : selection.anchor;

  let changed = true;
  while (changed) {
    changed = false;
    for (const control of controls) {
      if (!control.placeholder) continue;
      const range = placeholderSelectionRange(layout, control);
      if (!range) continue;
      const collapsed = comparePositions(layout, from, to) === 0;
      const intersects = collapsed
        ? comparePositions(layout, from, range.from) >= 0 &&
          comparePositions(layout, from, range.to) < 0
        : comparePositions(layout, from, range.to) < 0 &&
          comparePositions(layout, to, range.from) > 0;
      if (!intersects) continue;
      if (comparePositions(layout, range.from, from) < 0) {
        from = range.from;
        changed = true;
      }
      if (comparePositions(layout, range.to, to) > 0) {
        to = range.to;
        changed = true;
      }
    }
  }

  // Preserve which end the user is dragging (head) relative to the original orientation.
  const headIsEnd = comparePositions(layout, selection.anchor, selection.head) <= 0;
  return headIsEnd ? { anchor: from, head: to } : { anchor: to, head: from };
}
