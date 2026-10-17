// The host's say over what a local write publishes, asked at the store's commit boundary.

import type { OoxmlPackage } from '../package/ooxml-package.ts';
import type { OoxmlPart } from '../package/ooxml-tree.ts';
import { ORIGIN_IDS } from '../registry/frozen-ids.ts';

/**
 * Asked about every part a local write changes, as that part would be published, after the
 * engine's own invariants pass. A string refuses the whole write with that reason; `null`
 * admits it. Every lane that writes reaches it: typing, paste, drop, menus, automation and note
 * or header insertion. @public
 */
export type WriteAdmission = (request: { readonly part: OoxmlPart }) => string | null;

/**
 * The origins a host decides on. A remote write is another participant's admitted state, and
 * undo or redo restores state already admitted; refusing them would split the replicas.
 */
const ADMITTED_ORIGINS: ReadonlySet<string> = new Set([
  ORIGIN_IDS.mutationHuman,
  ORIGIN_IDS.mutationAgent,
]);

/**
 * The first refusal over the parts `after` changes from `before`, or null. The admission is read
 * only for an origin the host decides on.
 */
export function writeAdmissionRefusal(
  admissionOf: (() => WriteAdmission | undefined) | undefined,
  origin: string,
  before: OoxmlPackage,
  after: OoxmlPackage
): string | null {
  if (!ADMITTED_ORIGINS.has(origin)) return null;
  const admission = admissionOf?.();
  if (!admission) return null;
  for (const [name, part] of after.parts) {
    if (before.parts.get(name) === part) continue;
    const refusal = admission({ part });
    if (refusal !== null) return refusal;
  }
  return null;
}
