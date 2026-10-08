// What an untracked deletion does to the inline content controls it reaches.
//
// A range that holds a control and reaches past it takes the control with it, as Word does:
// the wrapper used to stay behind empty, a zero-width shell no caret could enter and no key
// could remove. A range inside one control only empties that control, which shows its prompt
// again (`content-control-prompt-restore.ts`). A wrapper the document locks against deletion
// (`sdtLocked`) stays, emptied, like the control the range lies in. A prompt is state, not text:
// a deletion that reaches into one leaves it whole, never a few of its letters.

import { contentControlPropertiesOf, contentControlsIn } from '../package/content-control-nodes.ts';
import type { EditOptions } from '../package/ooxml-edit.ts';
import type { OoxmlParagraphNode } from '../package/ooxml-tree.ts';
import { restoreEmptiedPlaceholder, rewritePrompt } from './content-control-prompt-restore.ts';
import { applyRemoveContentControl } from './tree-op-content-controls.ts';
import { paragraphModelTextOf } from './paragraph-model-text.ts';
import { paragraphOffsetIndex } from './tree-op-segments.ts';
import type { TreeOpResult } from './tree-op-types.ts';

/** One paragraph range an untracked deletion removed, with the paragraph as it was before. */
export interface DeletedRange {
  readonly paragraph: OoxmlParagraphNode;
  readonly start: number;
  readonly end: number;
}

interface DeletionCoverage {
  /** Every control the range holds whole and reaches past, outermost first. */
  readonly covered: readonly string[];
  /** The innermost control holding the whole range, or null when the range lies in none. */
  readonly container: string | null;
  /** The controls showing a prompt the range reaches into without covering them, and that prompt. */
  readonly prompts: readonly { readonly controlId: string; readonly text: string }[];
}

/** Read on the paragraph BEFORE the deletion: afterwards a covered control has no span left. */
function deletionCoverageOf({ paragraph, start, end }: DeletedRange): DeletionCoverage {
  const index = paragraphOffsetIndex(paragraph);
  const covered: string[] = [];
  const text = paragraphModelTextOf(paragraph);
  const prompts: { readonly controlId: string; readonly text: string }[] = [];
  let container: { readonly id: string; readonly depth: number } | null = null;
  for (const entry of contentControlsIn(paragraph)) {
    const span = index.spanOf(entry.node);
    if (!span) continue;
    const reached = span.start < end && span.end > start;
    if (reached && contentControlPropertiesOf(entry.node).showingPlaceholder) {
      prompts.push({ controlId: entry.node.id, text: text.slice(span.start, span.end) });
    }
    if (span.start <= start && span.end >= end) {
      if (!container || entry.depth > container.depth) {
        container = { id: entry.node.id, depth: entry.depth };
      }
      continue;
    }
    // A shell already empty is covered only strictly inside the range: one at its edge is
    // beside what was deleted, not in it.
    const holdsSomething = span.end > span.start || (start < span.start && span.start < end);
    if (span.start >= start && span.end <= end && holdsSomething) covered.push(entry.node.id);
  }
  return {
    covered,
    container: container?.id ?? null,
    prompts: prompts.filter((prompt) => !covered.includes(prompt.controlId)),
  };
}

/**
 * After the text went: each covered control leaves (unwrapped, so a bookmark or comment marker
 * it still held stays in the paragraph), a prompt the range reached into is written back whole,
 * and the control the range lay in shows its prompt if the deletion emptied it.
 */
export function settleDeletionCoverage(
  result: TreeOpResult,
  deleted: DeletedRange,
  options?: EditOptions
): TreeOpResult {
  const coverage = deletionCoverageOf(deleted);
  let settled = result;
  for (const controlId of coverage.covered) {
    if (!settled.ok) return settled;
    const removed = applyRemoveContentControl(
      settled.part,
      { op: 'removeContentControl', controlId, keepContent: true },
      options
    );
    if (removed.ok) {
      settled = {
        ok: true,
        part: removed.part,
        effect: {
          ...settled.effect,
          dirty: [...new Set([...settled.effect.dirty, ...removed.effect.dirty])],
          impact: 'flow-structural',
        },
      };
    } else if (removed.reason === 'locked') {
      settled = restoreEmptiedPlaceholder(settled, controlId, options);
    } else {
      return removed;
    }
  }
  for (const prompt of coverage.prompts) {
    settled = rewritePrompt(settled, prompt.controlId, options, prompt.text);
  }
  return coverage.container === null
    ? settled
    : restoreEmptiedPlaceholder(settled, coverage.container, options);
}
