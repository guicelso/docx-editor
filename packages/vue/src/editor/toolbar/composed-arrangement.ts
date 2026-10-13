// Where a registry slot that the host composes OUTSIDE the default arrangement goes.
//
// A part child for a slot the default bar does not draw (inserting a table, a footnote, the
// paragraph marks) used to append after the arrangement as host content, costed as fixed width:
// it never collapsed, and on a narrow window the row ran past the bar and under whatever the
// host put beside it. It is not host content. Its label, icon, command and GROUP are the
// registry's, so it takes its registry group: inside the default group that holds it, in
// registry order, or in a group of its own after the default ones, which collapses into "⋯"
// under the registry label like any other. Only children that drive no slot stay fixed.
//
// FRAMEWORK-NEUTRAL, byte-identical in both adapters: the entry type is the adapter's own.

import { CHROME_GROUPS, chromeSlotId, type ChromeSlotId } from '@docx-editor.dev/core/editor';

/** A group as the toolbar arranges it: the collapse unit, and a section of the "⋯" panel. */
export interface ArrangedGroup<Entry> {
  readonly id: string;
  readonly labelKey: string;
  readonly entries: readonly Entry[];
}

interface RegistryPlace {
  readonly groupIndex: number;
  readonly controlIndex: number;
}

const PLACES: ReadonlyMap<string, RegistryPlace> = new Map(
  CHROME_GROUPS.flatMap((group, groupIndex) =>
    group.controls.map(
      (control, controlIndex) =>
        [chromeSlotId(group, control), { groupIndex, controlIndex }] as const
    )
  )
);

/** Whether `slot` names a control of the chrome registry. */
export function isRegistrySlot(slot: string): slot is ChromeSlotId {
  return PLACES.has(slot);
}

/**
 * The default groups with every composed slot in its registry group: merged into the default
 * group that holds it, or in a group of its own after the default ones. Both keep registry order.
 */
export function arrangeComposedSlots<Entry extends { readonly slot: string }>(
  defaults: readonly ArrangedGroup<Entry>[],
  composed: readonly ChromeSlotId[],
  entryOf: (slot: ChromeSlotId) => Entry
): readonly ArrangedGroup<Entry>[] {
  if (composed.length === 0) return defaults;

  const byGroup = new Map<number, ChromeSlotId[]>();
  for (const slot of composed) {
    const place = PLACES.get(slot);
    if (place) byGroup.set(place.groupIndex, [...(byGroup.get(place.groupIndex) ?? []), slot]);
  }

  const defaultIds = new Set(defaults.map((group) => group.id));
  const merged = defaults.map((group) => {
    const extra = byGroup.get(CHROME_GROUPS.findIndex((candidate) => candidate.id === group.id));
    if (!extra) return group;
    return { ...group, entries: [...group.entries, ...extra.map(entryOf)].sort(byRegistryOrder) };
  });
  const own = [...byGroup.entries()]
    .filter(([groupIndex]) => !defaultIds.has(CHROME_GROUPS[groupIndex]!.id))
    .sort(([left], [right]) => left - right)
    .map(([groupIndex, slots]) => {
      const group = CHROME_GROUPS[groupIndex]!;
      return {
        id: group.id,
        labelKey: group.labelKey,
        entries: slots.map(entryOf).sort(byRegistryOrder),
      };
    });

  return [...merged, ...own];
}

function byRegistryOrder(
  left: { readonly slot: string },
  right: { readonly slot: string }
): number {
  return controlIndexOf(left.slot) - controlIndexOf(right.slot);
}

/** A key the registry does not name (the merged `'alignment'` dropdown) stays first in its group. */
function controlIndexOf(slot: string): number {
  return PLACES.get(slot)?.controlIndex ?? -1;
}
