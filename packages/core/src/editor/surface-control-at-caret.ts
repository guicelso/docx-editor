// The content control the caret is in: one answer, for the chrome the surface paints active and
// for the host reading `activeControlId`. Its own module because the surface sits at its line cap.

import {
  contentControlHoldingParagraph,
  contentControlRecordsInPart,
  contentControlsInLayout,
  sameCaretSlotNeighbour,
  type ContentControlBoundaryRecord,
  type SemanticLayout,
  type SemanticSelection,
  type TextMeasurer,
} from '@docx-editor.dev/core/layout';
import type { OoxmlPart, StoryScope } from '@docx-editor.dev/core/store';
import { contentControlAncestorsOf } from '../store/store/tree-op-nodes.ts';
import { contentControlAtSelection } from './content-control-at-selection.ts';
import { slotOwnerOf, type SlotNeighbours } from './surface-caret-slots.ts';

export interface ControlAtCaretDeps {
  readonly storyScope: () => StoryScope;
  readonly bodyPart: () => OoxmlPart;
  readonly storyPart: (scope: StoryScope) => OoxmlPart | null;
  readonly layout: () => SemanticLayout;
  readonly selection: () => SemanticSelection;
  readonly measurer: () => TextMeasurer;
  /** The slot the caret stands in where control edges meet, or null where the offset alone decides. */
  readonly slot: () => SlotNeighbours | null;
}

/**
 * Where control edges meet, the slot says which control the caret is in: one offset is several
 * places there, and only the slot is where typing lands. Elsewhere, two resolutions, because the
 * boundary records the fast one needs exist for the body alone. In the body, geometry: a paragraph
 * can hold several inline controls and only the caret's x/y says which. Anywhere else, the caret's
 * own PARAGRAPH, walked in that story's part — a header caret's y lands in the top band of the BODY
 * content box, so geometry there answered with a body control, and `setValue` and `remove` then
 * rewrote body content while the reader was editing a header.
 */
export function createControlAtCaret(
  deps: ControlAtCaretDeps
): () => ContentControlBoundaryRecord | null {
  // `state()` resolves the active control on every read, and hosts read state repeatedly between
  // changes. Layout and selection are replaced, never mutated, so identity is a sound key; the slot
  // moves on its own (an arrow between tags), so it is compared by its edges. The story branch stays
  // unmemoized: it reads the live part, which no key covers.
  let memo: {
    readonly layout: SemanticLayout;
    readonly selection: SemanticSelection;
    readonly slot: SlotNeighbours | null;
    readonly result: ContentControlBoundaryRecord | null;
  } | null = null;

  const inStory = (scope: StoryScope, slot: SlotNeighbours | null) => {
    const part = deps.storyPart(scope);
    if (!part) return null;
    if (!slot) return contentControlHoldingParagraph(part, deps.selection().head.paragraphId);
    return recordOf(contentControlRecordsInPart(part), slotOwnerOf(slot, parentIn(part)));
  };

  const inBody = (slot: SlotNeighbours | null) => {
    const layout = deps.layout();
    const part = deps.bodyPart();
    if (slot) return recordOf(contentControlsInLayout(layout), slotOwnerOf(slot, parentIn(part)));
    const found = contentControlAtSelection(layout, deps.selection(), deps.measurer());
    // The records are the body's, so a match from another part is a bug in the index rather than
    // an answer, and must not become a write.
    return found && partNameOfNodeId(found.id) === part.name ? found : null;
  };

  return () => {
    const scope = deps.storyScope();
    const slot = deps.slot();
    if (scope.kind !== 'body') return inStory(scope, slot);
    const layout = deps.layout();
    const selection = deps.selection();
    if (
      memo &&
      memo.layout === layout &&
      memo.selection === selection &&
      sameSlot(memo.slot, slot)
    ) {
      return memo.result;
    }
    const result = inBody(slot);
    memo = { layout, selection, slot, result };
    return result;
  };
}

function parentIn(part: OoxmlPart): (controlId: string) => string | null {
  return (controlId) => contentControlAncestorsOf(part, controlId).at(-1)?.id ?? null;
}

function recordOf(
  records: readonly ContentControlBoundaryRecord[],
  controlId: string | null
): ContentControlBoundaryRecord | null {
  return controlId === null ? null : (records.find((record) => record.id === controlId) ?? null);
}

function sameSlot(a: SlotNeighbours | null, b: SlotNeighbours | null): boolean {
  if (a === null || b === null) return a === b;
  return sameCaretSlotNeighbour(a.left, b.left) && sameCaretSlotNeighbour(a.right, b.right);
}

/** The part name a node id names, or '' when it names none. */
function partNameOfNodeId(id: string): string {
  const hash = id.indexOf('#');
  return hash === -1 ? '' : id.slice(0, hash);
}
