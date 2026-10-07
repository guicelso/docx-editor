// The caret slot the surface stands in at a tagged edge.
//
// With view-only content-control tags drawn, one model offset shows one slot per gap between
// chips, and the slot — not the offset — says where typing lands. This owns which slot the
// caret is in: the one a press or an arrow chose while the caret stays put, or the default.

import {
  caretSlotsAt,
  sameCaretSlotNeighbour,
  type CaretSlot,
  type CaretSlotNeighbour,
  type SemanticHitTag,
  type SemanticLayout,
  type SemanticPosition,
  type SemanticSelection,
} from '@docx-editor.dev/core/layout';
import type { ContentControlSurfaceState } from './surface-content-control-contract.ts';

/** Where text typed in a slot goes: beside a control, or at the start or end of one. */
export type SlotPlacement =
  | { readonly beside: { readonly controlId: string; readonly side: 'before' | 'after' } }
  | { readonly inside: string };

export interface CaretSlots {
  /** The slot the caret stands in, or null where no chip stands at a collapsed caret. */
  current(): CaretSlot | null;
  /** A plain arrow walks the slots first: true when it did, false when the offset must move. */
  step(direction: 'left' | 'right'): boolean;
  /** An arrow that moved the offset rightwards stops in the slot touching the text it left. */
  arrivedFromLeft(): void;
  /** The slot beside the chip a plain click landed on, on the half it landed on. */
  chooseBeside(tag: SemanticHitTag): void;
  /** A selection move forgets a chosen slot that is no longer where the caret stands. */
  forgetUnlessAt(next: SemanticSelection): void;
  /** Typing in front of a control stays in front of it, keystroke after keystroke. */
  keepInFront(typedIn: CaretSlot | null, landing: SemanticPosition): void;
  /** What a host reads of the slot: its two neighbours. */
  neighbours(): ContentControlSurfaceState['caretSlot'];
}

export function createCaretSlots(deps: {
  /** Whether tags are drawn at all; without them there is no slot. */
  readonly enabled: () => boolean;
  readonly selection: () => SemanticSelection;
  readonly layout: () => SemanticLayout;
  /** The slot moved: the caret repaints and the state is published. */
  readonly changed: () => void;
}): CaretSlots {
  let chosen: {
    readonly at: SemanticPosition;
    readonly left: CaretSlotNeighbour | null;
    readonly right: CaretSlotNeighbour | null;
  } | null = null;

  const all = (): readonly CaretSlot[] => {
    const selection = deps.selection();
    return deps.enabled() && isCollapsedSelection(selection)
      ? caretSlotsAt(deps.layout(), selection.head)
      : [];
  };

  /**
   * A chosen slot holds while the caret stays put; otherwise the slot touching the text on the
   * RIGHT — where a programmatic caret, an undo or a Home/End lands.
   */
  const currentOf = (slots: readonly CaretSlot[]): CaretSlot | null => {
    const held = chosen;
    const kept =
      held && samePosition(held.at, deps.selection().head)
        ? slots.find(
            (slot) =>
              sameCaretSlotNeighbour(slot.left, held.left) &&
              sameCaretSlotNeighbour(slot.right, held.right)
          )
        : undefined;
    return kept ?? slots[slots.length - 1] ?? null;
  };

  const choose = (slot: CaretSlot): void => {
    chosen = { at: deps.selection().head, left: slot.left, right: slot.right };
    deps.changed();
  };

  return {
    current: () => currentOf(all()),
    step(direction) {
      const slots = all();
      const current = currentOf(slots);
      const next = current
        ? slots[slots.indexOf(current) + (direction === 'right' ? 1 : -1)]
        : undefined;
      if (next) choose(next);
      return next !== undefined;
    },
    arrivedFromLeft() {
      const [first] = all();
      if (first) choose(first);
    },
    chooseBeside(tag) {
      const slot = all().find((candidate) =>
        sameCaretSlotNeighbour(tag.side === 'before' ? candidate.right : candidate.left, tag)
      );
      if (slot) choose(slot);
    },
    forgetUnlessAt(next) {
      if (chosen && !(isCollapsedSelection(next) && samePosition(next.head, chosen.at))) {
        chosen = null;
      }
    },
    keepInFront(typedIn, landing) {
      if (typedIn?.right?.edge !== 'open' || !samePosition(deps.selection().head, landing)) return;
      chosen = { at: landing, left: null, right: typedIn.right };
      deps.changed();
    },
    neighbours() {
      const slot = currentOf(all());
      return slot ? { left: slot.left, right: slot.right } : null;
    },
  };
}

/**
 * An opening tag on the right means in front of that control; a closing one on the left,
 * behind it; otherwise the start or the end of the control whose tag is beside the slot.
 */
export function slotPlacementOf(slot: CaretSlot): SlotPlacement {
  if (slot.right?.edge === 'open') {
    return { beside: { controlId: slot.right.controlId, side: 'before' } };
  }
  if (slot.left?.edge === 'close') {
    return { beside: { controlId: slot.left.controlId, side: 'after' } };
  }
  return { inside: (slot.left ?? slot.right)!.controlId };
}

export function isCollapsedSelection(selection: SemanticSelection): boolean {
  return samePosition(selection.anchor, selection.head);
}

function samePosition(a: SemanticPosition, b: SemanticPosition): boolean {
  return a.paragraphId === b.paragraphId && a.offset === b.offset;
}
