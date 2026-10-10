// The caret slots between view-only content-control tags.
//
// A tagged edge paints several chips over ONE model offset, so one offset shows several places
// a caret can stand: in front of a control, at the start of its content, at the end, behind
// it. A slot is named by its two NEIGHBOURS — the chip on its left and the chip on its right —
// because "the control the caret is in" is exactly what is ambiguous there.

import type { ContentControlTagMark } from './content-control-tags.ts';
import { paragraphLinesIndex } from './paragraph-lines.ts';
import type { SemanticPosition } from './semantic-interaction.ts';
import type { SemanticLayout } from './semantic-records.ts';

/** One chip beside a slot: which control, and which of its edges. @public */
export interface CaretSlotNeighbour {
  readonly controlId: string;
  readonly edge: ContentControlTagMark['edge'];
  readonly level: ContentControlTagMark['level'];
}

/** Where a slot stands: its neighbours, and its place on the laid-out line. @public */
export interface CaretSlot {
  readonly left: CaretSlotNeighbour | null;
  readonly right: CaretSlotNeighbour | null;
  readonly pageIndex: number;
  readonly lineId: string;
  /** Page-relative, in the coordinate space of the line boxes, as `CaretGeometry`. */
  readonly x: number;
  readonly y: number;
  readonly height: number;
}

interface PlacedChip {
  readonly mark: CaretSlotNeighbour;
  readonly pageIndex: number;
  readonly lineId: string;
  readonly left: number;
  readonly right: number;
  readonly y: number;
  readonly height: number;
}

/**
 * Every slot at a position, in reading order, or none where no chip stands. Slot `0` touches
 * the text on the left and the last one the text on the right.
 */
export function caretSlotsAt(
  layout: SemanticLayout,
  position: SemanticPosition
): readonly CaretSlot[] {
  const chips = chipsAt(layout, position);
  if (chips.length === 0) return [];
  const slots: CaretSlot[] = [];
  for (let index = 0; index <= chips.length; index += 1) {
    const before = chips[index - 1];
    const after = chips[index];
    // The slot stands on the chip it touches: the one after it, or the last one.
    const host = after ?? before!;
    slots.push({
      left: before?.mark ?? null,
      right: after?.mark ?? null,
      pageIndex: host.pageIndex,
      lineId: host.lineId,
      x: after ? after.left : before!.right,
      y: host.y,
      height: host.height,
    });
  }
  return slots;
}

/** Whether two neighbours name the same chip: one control, one edge. */
export function sameCaretSlotNeighbour(
  a: Pick<CaretSlotNeighbour, 'controlId' | 'edge'> | null,
  b: Pick<CaretSlotNeighbour, 'controlId' | 'edge'> | null
): boolean {
  return a === b || (a !== null && b !== null && a.controlId === b.controlId && a.edge === b.edge);
}

function chipsAt(layout: SemanticLayout, position: SemanticPosition): readonly PlacedChip[] {
  const chips: PlacedChip[] = [];
  for (const { line, pageIndex } of paragraphLinesIndex(layout).get(position.paragraphId) ?? []) {
    for (const span of line.spans) {
      const tag = span.contentControlTag;
      if (!tag || span.range.paragraphId !== position.paragraphId) continue;
      if (span.range.start !== position.offset) continue;
      chips.push({
        mark: { controlId: tag.controlId, edge: tag.edge, level: tag.level },
        pageIndex,
        lineId: line.id,
        left: span.box.x,
        right: span.box.x + span.box.width,
        y: line.box.y,
        height: line.box.height,
      });
    }
  }
  return chips;
}
