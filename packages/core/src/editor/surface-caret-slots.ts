// The caret slot the surface stands in where inline content-control edges meet.
//
// One model offset where controls start or end is several places: in front of a control, at the
// start of its content, at its end, behind it. A slot is named by the edge on each side of it,
// because "the control the caret is in" is exactly what is ambiguous there. With view-only tags
// drawn every slot is visible, and a press or an arrow chooses one. Without them an edit still
// knows the slot it leaves the caret in, so the next keystroke continues where the last one went.
// An atom is one unit, its tags with it: no slot stands inside it, and a slot beside one of its tags
// stands outside it, before or after the whole atom.

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
import type { BlockPlace } from '../store/store/tree-op-block-place.ts';
import type { ContentControlSurfaceState } from './surface-content-control-contract.ts';

/**
 * Where text typed in a slot goes: beside an inline control, at the start or end of one, or — in
 * the slot outside a block control's tag — at the place between blocks before or after it. @public
 */
export type SlotPlacement =
  | { readonly beside: { readonly controlId: string; readonly side: 'before' | 'after' } }
  | { readonly inside: string }
  | { readonly block: BlockPlace };

/** The edge on each side of a slot; `null` is text. */
export interface SlotNeighbours {
  readonly left: CaretSlotNeighbour | null;
  readonly right: CaretSlotNeighbour | null;
}

/**
 * The side of its slot an edit keeps: typing and Backspace keep what stood on the caret's right,
 * Delete what stood on its left. The other side is whatever the edit left there.
 */
export type SlotKeep =
  | { readonly left: CaretSlotNeighbour | null }
  | { readonly right: CaretSlotNeighbour | null };

/** Typed text stands on the caret's left: the caret touches what it typed. */
export const AFTER_TYPED_TEXT: SlotKeep = { left: null };

export interface CaretSlots {
  /** The drawn slot the caret stands in, or null where no tag stands at a collapsed caret. */
  current(): CaretSlot | null;
  /**
   * Where text at the collapsed caret lands, or null where the offset alone decides: no edge
   * meets there, or tags are hidden and no edit has placed the caret.
   */
  placement(): SlotPlacement | null;
  /** The side of the caret's slot an edit at the caret keeps, or undefined where none is known. */
  keep(side: 'left' | 'right'): SlotKeep | undefined;
  /** What removing a range that starts at `position` keeps: the last edge in front of it. */
  keepFrontOf(position: SemanticPosition): SlotKeep;
  /** A plain arrow walks the drawn slots first: true when it did, false when the offset must move. */
  step(direction: 'left' | 'right'): boolean;
  /** An arrow that moved the offset rightwards stops in the slot touching the text it left. */
  arrivedFromLeft(): void;
  /** The slot beside a tag, on the side named: a click on a drawn tag, or a host placing the caret. */
  chooseBeside(tag: SemanticHitTag): void;
  /** A selection move forgets a slot that is no longer where the caret stands. */
  forgetUnlessAt(next: SemanticSelection): void;
  /** An edit left the caret at `at`, in the slot keeping one side of the one it started from. */
  hold(at: SemanticPosition, keep: SlotKeep): void;
  /** The caret where it stands takes the slot keeping that side, as a host asked. */
  stand(keep: SlotKeep): void;
  /** What a host reads of the slot the caret stands in: its two neighbours and where a write lands, where one is known. */
  neighbours(): ContentControlSurfaceState['caretSlot'];
}

/** A slot held at a position; a side left out is whatever stands there. */
interface HeldSlot {
  readonly at: SemanticPosition;
  readonly left?: CaretSlotNeighbour | null;
  readonly right?: CaretSlotNeighbour | null;
}

export function createCaretSlots(deps: {
  /** Whether tags are drawn: only drawn slots are walked, clicked and painted. */
  readonly enabled: () => boolean;
  readonly selection: () => SemanticSelection;
  readonly layout: () => SemanticLayout;
  /** The control edges at a position, in reading order, from the document itself. */
  readonly edgesAt: (position: SemanticPosition) => readonly CaretSlotNeighbour[];
  /** The controls the host made atoms, as layout published them. */
  readonly atomIds: () => ReadonlySet<string>;
  /** The slot moved: the caret repaints and the state is published. */
  readonly changed: () => void;
}): CaretSlots {
  let held: HeldSlot | null = null;

  const collapsedHead = (): SemanticPosition | null => {
    const selection = deps.selection();
    return isCollapsedSelection(selection) ? selection.head : null;
  };

  const drawnAt = (position: SemanticPosition): readonly CaretSlot[] =>
    deps.enabled()
      ? outsideAtoms(caretSlotsAt(deps.layout(), position), deps.edgesAt(position), deps.atomIds())
      : [];

  /** Every slot at a position: the drawn ones, or the document's own edges with tags hidden. */
  const slotsAt = (position: SemanticPosition): readonly SlotNeighbours[] => {
    if (deps.enabled()) return drawnAt(position);
    const edges = deps.edgesAt(position);
    return outsideAtoms(slotsBetween(edges), edges, deps.atomIds());
  };

  const heldAmong = <Slot extends SlotNeighbours>(slots: readonly Slot[]): Slot | undefined => {
    const head = collapsedHead();
    const hold = held;
    if (!hold || !head || !samePosition(hold.at, head)) return undefined;
    return slots.find(
      (slot) =>
        (hold.left === undefined || sameCaretSlotNeighbour(slot.left, hold.left)) &&
        (hold.right === undefined || sameCaretSlotNeighbour(slot.right, hold.right))
    );
  };

  /**
   * A held slot stands while the caret stays put; otherwise the drawn slot touching the text on
   * the RIGHT — where a programmatic caret, an undo or a Home/End lands.
   */
  const currentAmong = (slots: readonly CaretSlot[]): CaretSlot | null =>
    heldAmong(slots) ?? slots[slots.length - 1] ?? null;

  const drawnHere = (): readonly CaretSlot[] => {
    const head = collapsedHead();
    return head ? drawnAt(head) : [];
  };

  const current = (): CaretSlot | null => currentAmong(drawnHere());

  /**
   * Where no slot is drawn, an atom's edge still decides: the caret there stands outside the atom
   * — in the slot an edit held, else the one touching the text on the right — and never in it.
   */
  const besideAtom = (head: SemanticPosition): SlotNeighbours | null => {
    const edges = deps.edgesAt(head);
    const atomIds = deps.atomIds();
    if (!edges.some((edge) => atomIds.has(edge.controlId))) return null;
    const slots = outsideAtoms(slotsBetween(edges), edges, atomIds);
    return heldAmong(slots) ?? slots[slots.length - 1] ?? null;
  };

  const place = (): SlotNeighbours | null => {
    const head = collapsedHead();
    if (!head) return null;
    if (deps.enabled()) return current() ?? besideAtom(head) ?? TEXT_BOTH_SIDES;
    const slots = slotsAt(head);
    if (slots.length === 0) return TEXT_BOTH_SIDES;
    return heldAmong(slots) ?? besideAtom(head);
  };

  const choose = (slot: SlotNeighbours): void => {
    const head = collapsedHead();
    if (!head) return;
    held = { at: head, left: slot.left, right: slot.right };
    deps.changed();
  };

  return {
    current,
    placement() {
      const slot = place();
      return slot && (slot.left || slot.right) ? slotPlacementOf(slot, deps.atomIds()) : null;
    },
    keep(side) {
      const slot = place();
      if (!slot) return undefined;
      return side === 'left' ? { left: slot.left } : { right: slot.right };
    },
    keepFrontOf(position) {
      const slots = slotsAt(position);
      return { left: slots[slots.length - 1]?.left ?? null };
    },
    step(direction) {
      const slots = drawnHere();
      const at = currentAmong(slots);
      const next = at ? slots[slots.indexOf(at) + (direction === 'right' ? 1 : -1)] : undefined;
      if (next) choose(next);
      return next !== undefined;
    },
    arrivedFromLeft() {
      const [first] = drawnHere();
      if (first) choose(first);
    },
    chooseBeside(tag) {
      const head = collapsedHead();
      const slot = (head ? slotsAt(head) : []).find((candidate) =>
        sameCaretSlotNeighbour(tag.side === 'before' ? candidate.right : candidate.left, tag)
      );
      if (slot) choose(slot);
    },
    forgetUnlessAt(next) {
      if (held && !(isCollapsedSelection(next) && samePosition(next.head, held.at))) held = null;
    },
    hold(at, keep) {
      held = { at, ...keep };
    },
    stand(keep) {
      const head = collapsedHead();
      if (!head) return;
      held = { at: head, ...keep };
      deps.changed();
    },
    neighbours() {
      const slot = place();
      return slot && (slot.left || slot.right)
        ? { left: slot.left, right: slot.right, placement: slotPlacementOf(slot, deps.atomIds()) }
        : null;
    },
  };
}

/** Text on both sides of the caret: the place is the offset, whatever the edit. */
const TEXT_BOTH_SIDES: SlotNeighbours = { left: null, right: null };

/** The slots between edges in reading order: one before the first, one after each. */
function slotsBetween(edges: readonly CaretSlotNeighbour[]): readonly SlotNeighbours[] {
  if (edges.length === 0) return [];
  return [null, ...edges].map((left, index) => ({ left, right: edges[index] ?? null }));
}

/**
 * The slots that stand outside every atom. A slot spans the edges between its two neighbours in
 * the document's reading order; it stands inside an atom when that span lies between the atom's
 * opening and closing edges — an edge not at this offset lies beyond the slot on its side, as the
 * atom's content does.
 */
function outsideAtoms<Slot extends SlotNeighbours>(
  slots: readonly Slot[],
  edges: readonly CaretSlotNeighbour[],
  atomIds: ReadonlySet<string>
): readonly Slot[] {
  const atoms = atomSpansIn(edges, atomIds);
  if (atoms.length === 0) return slots;
  return slots.filter((slot) => {
    const left = slot.left === null ? -1 : edgeIndexOf(edges, slot.left);
    const right = slot.right === null ? edges.length : edgeIndexOf(edges, slot.right);
    if (left === undefined || right === undefined) return true;
    return !atoms.some((atom) => left >= atom.open && right <= atom.close);
  });
}

/** Where each atom with an edge at this offset opens and closes, among the offset's edges. */
function atomSpansIn(
  edges: readonly CaretSlotNeighbour[],
  atomIds: ReadonlySet<string>
): readonly { readonly open: number; readonly close: number }[] {
  const ids = new Set(edges.map((edge) => edge.controlId).filter((id) => atomIds.has(id)));
  return [...ids].map((controlId) => {
    const open = edges.findIndex((edge) => edge.controlId === controlId && edge.edge === 'open');
    const close = edges.findIndex((edge) => edge.controlId === controlId && edge.edge === 'close');
    return { open, close: close === -1 ? edges.length : close };
  });
}

function edgeIndexOf(
  edges: readonly CaretSlotNeighbour[],
  neighbour: CaretSlotNeighbour
): number | undefined {
  const index = edges.findIndex((edge) => sameCaretSlotNeighbour(edge, neighbour));
  return index === -1 ? undefined : index;
}

/**
 * A slot beside an atom's edge stands outside the atom: in front of it on the atom's left, behind
 * it on its right. Otherwise an opening edge on the right means in front of that control; a
 * closing one on the left, behind it; otherwise the start or the end of the control whose edge is
 * beside the slot. The start or end of a BLOCK control names no place: its first or last paragraph
 * is already the place, and the offset says where in it, as it does with text on both sides.
 */
export function slotPlacementOf(
  slot: SlotNeighbours,
  atomIds: ReadonlySet<string>
): SlotPlacement | null {
  if (slot.right && (slot.right.edge === 'open' || atomIds.has(slot.right.controlId))) {
    const controlId = slot.right.controlId;
    return slot.right.level === 'block'
      ? { block: { before: controlId } }
      : { beside: { controlId, side: 'before' } };
  }
  if (slot.left && (slot.left.edge === 'close' || atomIds.has(slot.left.controlId))) {
    const controlId = slot.left.controlId;
    return slot.left.level === 'block'
      ? { block: { after: controlId } }
      : { beside: { controlId, side: 'after' } };
  }
  const owner = slot.left ?? slot.right;
  return owner === null || owner.level === 'block' ? null : { inside: owner.controlId };
}

/**
 * The control a slot stands in. In front of an opening edge or behind a closing one — or beside
 * any edge of an atom — the caret is outside that control, in the control around it; past an
 * opening edge or before a closing one, it is inside. The same reading `slotPlacementOf` gives a
 * write, so the active control is where typing lands.
 */
export function slotOwnerOf(
  slot: SlotNeighbours,
  parentOf: (controlId: string) => string | null,
  atomIds: ReadonlySet<string>
): string | null {
  if (slot.right && (slot.right.edge === 'open' || atomIds.has(slot.right.controlId))) {
    return parentOf(slot.right.controlId);
  }
  if (slot.left && (slot.left.edge === 'close' || atomIds.has(slot.left.controlId))) {
    return parentOf(slot.left.controlId);
  }
  return (slot.left ?? slot.right)?.controlId ?? null;
}

export function isCollapsedSelection(selection: SemanticSelection): boolean {
  return samePosition(selection.anchor, selection.head);
}

function samePosition(a: SemanticPosition, b: SemanticPosition): boolean {
  return a.paragraphId === b.paragraphId && a.offset === b.offset;
}
