// Content inserted at a field's edge never joins the run that carries the field's chrome.
//
// `w:fldChar` and `w:instrText` belong to the field; what a person types or places beside it is
// theirs. CT_R lets one `w:r` hold both (§17.3.2.25), so joining the run that opens or closes the
// field is legal — and a consumer that replaces a field by its runs then takes the typed text with
// it: a merge that fills `MERGEFIELD` by swapping begin…end lost the comma typed after every field.
// The content goes in a run of its own beside that run, with a copy of its properties, so the
// formatting the caret showed is the one the content keeps, exactly as joining would have given.
//
// Only the run's EDGES move: an index between the run's own content (a run some other producer
// already wrote mixed) keeps the insertion where it is, because a run beside it would land on the
// wrong side of the content before the index.

import { isFieldChrome } from '../package/field-nodes.ts';
import type { OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { cloneWithNewIds, isRunPropertiesNode, parentOf } from './tree-op-nodes.ts';

/** Where a run of its own goes, and the properties it copies from the field's run. */
export interface FieldChromeLanding {
  readonly holderId: string;
  readonly index: number;
  readonly properties: readonly OoxmlNode[];
}

/**
 * The landing beside `run` for content that would be inserted at child `index` of it, or `null`
 * when `run` carries no field chrome or `index` is not one of its edges.
 */
export function landingBesideFieldChrome(
  part: OoxmlPart,
  run: OoxmlNode,
  index: number,
  nextId: () => string
): FieldChromeLanding | null {
  if (run.kind !== 'run' || !run.children.some(isFieldChrome)) return null;
  const side =
    index >= run.children.length
      ? 'after'
      : run.children.slice(0, index).every(isRunPropertiesNode)
        ? 'before'
        : null;
  if (side === null) return null;
  const holder = parentOf(part, run.id);
  if (!holder) return null;
  const at = holder.children.findIndex((child) => child.id === run.id);
  if (at < 0) return null;
  return {
    holderId: holder.id,
    index: side === 'after' ? at + 1 : at,
    properties: run.children
      .filter((child) => isRunPropertiesNode(child))
      .map((child) => cloneWithNewIds(child, nextId)),
  };
}
