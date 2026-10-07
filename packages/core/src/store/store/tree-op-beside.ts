// Where text lands BESIDE an inline content control: as its sibling, at its edge.

import type { OoxmlParagraphNode } from '../package/ooxml-tree.ts';
import { directParentOf, paragraphOffsetIndex, type InsertionSite } from './tree-op-segments.ts';

/**
 * The sibling slot right before or after an inline control, or null when the control is not in
 * this paragraph or `offset` is not that edge — a stale slot must not land somewhere else.
 */
export function siteBesideControl(
  paragraph: OoxmlParagraphNode,
  offset: number,
  beside: { readonly controlId: string; readonly side: 'before' | 'after' }
): InsertionSite | null {
  const holder = directParentOf(paragraph, beside.controlId);
  if (!holder) return null;
  const index = holder.children.findIndex((child) => child.id === beside.controlId);
  const control = holder.children[index];
  if (!control || control.kind === 'textValue') return null;
  const span = paragraphOffsetIndex(paragraph).spanOf(control);
  if (!span || (beside.side === 'before' ? span.start : span.end) !== offset) return null;
  return { kind: 'newRun', holder, index: beside.side === 'before' ? index : index + 1 };
}
