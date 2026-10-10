// What Backspace and Delete do where the caret meets an inline content control.
//
// With view-only tags drawn, a key whose neighbour in the caret's slot is a tag acts on that
// tag's control, never on a character on the far side of it: the first press selects the
// control whole, the next removes it with what it holds. Word selects a field before deleting
// from its boundary (`surface-text-form-fields.ts`); this is the same rule for a control, whose
// boundary the tags make visible. A control showing its prompt holds nothing to lose, so with
// the tags drawn a key in it removes it at once. Without drawn tags a key at an edge keeps Word's
// rule and takes one character. A wrapper locked against deletion is refused by the store, as
// every removal is.

import type {
  CaretSlot,
  CaretSlotNeighbour,
  SemanticPosition,
  SemanticSelection,
} from '@docx-editor.dev/core/layout';
import {
  contentControlLevelOf,
  contentControlPropertiesOf,
  contentControlsIn,
} from '../store/package/content-control-nodes.ts';
import { findNode } from '../store/package/ooxml-edit.ts';
import type { OoxmlNode, OoxmlPart } from '../store/package/ooxml-tree.ts';
import { paragraphIdsUnder } from '../store/store/content-control-value-content.ts';
import {
  paragraphIdsInDocumentOrder,
  survivingCaretAfterBlockRemoval,
} from '../store/store/tree-op-blocks.ts';
import { enclosingContentControls } from '../store/store/tree-op-content-controls.ts';
import { paragraphOffsetIndex } from '../store/store/tree-op-segments.ts';
import { selectionsEqual } from './dom-selection.ts';
import { isCollapsedSelection, type SlotKeep, type SlotPlacement } from './surface-caret-slots.ts';

export interface ControlKeyDeps {
  /** The story part that holds a paragraph. */
  readonly part: (paragraphId: string) => OoxmlPart;
  readonly selection: () => SemanticSelection;
  /** Whether view-only tags are drawn: only then does a key act on a control whole. */
  readonly tagsDrawn: () => boolean;
  /** The drawn slot the caret stands in, or null where no tag is drawn beside it. */
  readonly drawnSlot: () => CaretSlot | null;
  /** Where text at the caret lands, or null where the offset alone decides. */
  readonly placement: () => SlotPlacement | null;
  /** The control edges at a position, in reading order, from the document itself. */
  readonly edgesAt: (position: SemanticPosition) => readonly CaretSlotNeighbour[];
  readonly select: (selection: SemanticSelection) => void;
  /** Remove a control and its content, leaving the caret where it stood, in the slot kept. */
  readonly remove: (controlId: string, caret: SemanticPosition, keep: SlotKeep) => void;
}

export interface ControlKeys {
  /** Backspace (`backward`) or Delete (`forward`) at the selection: true when it took a control. */
  press(direction: 'backward' | 'forward'): boolean;
  /** Remove the control the selection holds whole: true when there was one. */
  removeSelected(): boolean;
  /** The control selected whole, while the selection still selects it. */
  selectedId(): string | null;
  /** A selection move forgets the selected control unless it still selects it. */
  forgetUnlessAt(next: SemanticSelection): void;
}

/** One control, with the positions its content spans: inside one paragraph, or across several. */
interface PlacedControl {
  readonly controlId: string;
  readonly from: SemanticPosition;
  readonly to: SemanticPosition;
  readonly depth: number;
  readonly prompt: boolean;
  readonly level: 'inline' | 'block';
}

export function createControlKeys(deps: ControlKeyDeps): ControlKeys {
  let selected: { readonly controlId: string; readonly selection: SemanticSelection } | null = null;

  const inlineControlsOf = (paragraphId: string): readonly PlacedControl[] => {
    const paragraph = findNode(deps.part(paragraphId), paragraphId);
    if (paragraph?.kind !== 'paragraph') return [];
    const index = paragraphOffsetIndex(paragraph);
    return contentControlsIn(paragraph).flatMap((entry): PlacedControl[] => {
      const span = index.spanOf(entry.node);
      if (!span) return [];
      return [
        {
          controlId: entry.node.id,
          from: { paragraphId, offset: span.start },
          to: { paragraphId, offset: span.end },
          depth: entry.depth,
          prompt: contentControlPropertiesOf(entry.node).showingPlaceholder,
          level: 'inline',
        },
      ];
    });
  };

  /** A block control spans its first paragraph's start to its last paragraph's end. */
  const blockControl = (part: OoxmlPart, control: OoxmlNode): PlacedControl | undefined => {
    if (control.kind !== 'contentControl' || contentControlLevelOf(control) !== 'block') return;
    const paragraphs = paragraphIdsUnder(control);
    const last = paragraphs[paragraphs.length - 1];
    const lastParagraph = last === undefined ? null : findNode(part, last);
    if (lastParagraph?.kind !== 'paragraph') return;
    return {
      controlId: control.id,
      from: { paragraphId: paragraphs[0]!, offset: 0 },
      to: { paragraphId: last!, offset: paragraphOffsetIndex(lastParagraph).length },
      depth: enclosingContentControls(part, control.id).length,
      prompt: contentControlPropertiesOf(control).showingPlaceholder,
      level: 'block',
    };
  };

  const placedById = (paragraphId: string, controlId: string): PlacedControl | undefined => {
    const inline = inlineControlsOf(paragraphId).find((control) => control.controlId === controlId);
    if (inline) return inline;
    const part = deps.part(paragraphId);
    const node = findNode(part, controlId);
    return node ? blockControl(part, node) : undefined;
  };

  const selectedNow = (): PlacedControl | null => {
    if (!selected || !selectionsEqual(deps.selection(), selected.selection)) return null;
    return placedById(deps.selection().head.paragraphId, selected.controlId) ?? null;
  };

  /** The caret stays where the control stood, in the slot touching the edge before its opening. */
  const removeWhole = (control: PlacedControl): void => {
    selected = null;
    if (control.level === 'block') {
      removeBlock(control);
      return;
    }
    const at = control.from;
    const edges = deps.edgesAt(at);
    const opening = edges.findIndex(
      (edge) => edge.controlId === control.controlId && edge.edge === 'open'
    );
    deps.remove(control.controlId, at, { left: opening > 0 ? edges[opening - 1]! : null });
  };

  /**
   * A block control goes with its paragraphs, so the caret goes to the paragraph that survives
   * nearest it — the end of the one before, else the start of the one after — outside every edge
   * there.
   */
  const removeBlock = (control: PlacedControl): void => {
    const part = deps.part(control.from.paragraphId);
    const survivor = survivingCaretAfterBlockRemoval(part, control.controlId);
    const paragraph = survivor === null ? null : findNode(part, survivor);
    if (paragraph?.kind !== 'paragraph') return;
    const order = paragraphIdsInDocumentOrder(part);
    const before = order.indexOf(paragraph.id) < order.indexOf(control.from.paragraphId);
    const at = {
      paragraphId: paragraph.id,
      offset: before ? paragraphOffsetIndex(paragraph).length : 0,
    };
    const edges = deps.edgesAt(at);
    deps.remove(
      control.controlId,
      at,
      before ? { left: edges[edges.length - 1] ?? null } : { right: edges[0] ?? null }
    );
  };

  /** The prompt the key is in: the control whose prompt holds the selection, innermost first. */
  const promptAtSelection = (): PlacedControl | undefined => {
    const { anchor, head } = deps.selection();
    if (anchor.paragraphId !== head.paragraphId) return undefined;
    const from = Math.min(anchor.offset, head.offset);
    const to = Math.max(anchor.offset, head.offset);
    const placement = isCollapsedSelection(deps.selection()) ? deps.placement() : null;
    if (placement && ('beside' in placement || 'block' in placement)) return undefined;
    const owner = inlineControlsOf(head.paragraphId)
      .filter((control) =>
        placement
          ? control.controlId === placement.inside
          : control.from.offset <= from && control.to.offset >= to
      )
      .reduce<PlacedControl | undefined>(
        (deepest, control) => (deepest && deepest.depth >= control.depth ? deepest : control),
        undefined
      );
    if (owner) return owner.prompt ? owner : undefined;
    return blockPromptAround(head.paragraphId);
  };

  /** The innermost block control holding the paragraph, when it shows its prompt. */
  const blockPromptAround = (paragraphId: string): PlacedControl | undefined => {
    const part = deps.part(paragraphId);
    const holders = enclosingContentControls(part, paragraphId);
    const innermost = holders[holders.length - 1];
    const control = innermost ? blockControl(part, innermost) : undefined;
    return control?.prompt ? control : undefined;
  };

  const removeSelected = (): boolean => {
    const control = selectedNow();
    if (!control) return false;
    removeWhole(control);
    return true;
  };

  return {
    press(direction) {
      if (removeSelected()) return true;
      if (!deps.tagsDrawn()) return false;
      const prompt = promptAtSelection();
      if (prompt) {
        removeWhole(prompt);
        return true;
      }
      const slot = isCollapsedSelection(deps.selection()) ? deps.drawnSlot() : null;
      const tag = direction === 'backward' ? slot?.left : slot?.right;
      const control = tag
        ? placedById(deps.selection().head.paragraphId, tag.controlId)
        : undefined;
      if (!control) return false;
      // Held BEFORE the selection moves, so the paint the move triggers draws it selected.
      selected = {
        controlId: control.controlId,
        selection: { anchor: control.from, head: control.to },
      };
      deps.select(selected.selection);
      return true;
    },
    removeSelected,
    selectedId: () => selectedNow()?.controlId ?? null,
    forgetUnlessAt(next) {
      if (selected && !selectionsEqual(next, selected.selection)) selected = null;
    },
  };
}
