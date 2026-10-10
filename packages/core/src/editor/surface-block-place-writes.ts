// What Enter and typing do in the slot outside a block control's tag: a new paragraph at the
// place between blocks, holding what was typed, with the caret in it.
//
// The paragraph's id is minted inside the transaction, so the caret is the one the op publishes
// (`TreeOpEffect.caret`), adopted from the committed change — the path `insertTable` takes.

import type { SemanticPosition, SemanticSelection } from '@docx-editor.dev/core/layout';
import type { TreeDocOp } from '@docx-editor.dev/core/store';
import type { TreeApplyResult } from '../binding/tree-session-contract.ts';
import type { BlockPlace } from '../store/store/tree-op-block-place.ts';
import { AFTER_TYPED_TEXT, type SlotKeep } from './surface-caret-slots.ts';

export interface BlockPlaceWriteDeps {
  /** Hear the committed change, for the caret the transaction published. */
  readonly subscribe: (
    listener: (change: { readonly caret?: { readonly paragraphId: string } | null }) => void
  ) => () => void;
  readonly commit: (
    run: () => TreeApplyResult,
    selectionAfter: () => SemanticSelection | null,
    options: { readonly slot: SlotKeep }
  ) => void;
  readonly applyOps: (ops: readonly TreeDocOp[]) => TreeApplyResult;
  readonly collapsedAt: (position: SemanticPosition) => SemanticSelection;
}

/** Write a paragraph at `at`, holding `text` when given, and leave the caret after it. */
export function createBlockPlaceWrites(
  deps: BlockPlaceWriteDeps
): (at: BlockPlace, text?: string) => void {
  return (at, text) => {
    let paragraphId: string | null = null;
    const unsubscribe = deps.subscribe((change) => {
      if (change.caret) paragraphId = change.caret.paragraphId;
    });
    try {
      deps.commit(
        () => deps.applyOps([{ op: 'insertParagraph', at, ...(text ? { text } : {}) }]),
        () =>
          paragraphId === null
            ? null
            : deps.collapsedAt({ paragraphId, offset: text?.length ?? 0 }),
        { slot: AFTER_TYPED_TEXT }
      );
    } finally {
      unsubscribe();
    }
  };
}
