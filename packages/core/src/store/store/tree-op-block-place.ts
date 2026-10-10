// A place BETWEEN blocks: before or after a paragraph, a table or a block-level control, or inside
// a block-level control that shows its prompt — the place a new block is written.
//
// The inline destinations (`inside` / `beside` in `tree-op-inline-destination.ts`) name places
// within a paragraph. A block place is the other level: what stands there is a whole block, so
// it is resolved against the anchor's container rather than against offsets.

import {
  contentControlLevelOf,
  contentControlPropertiesOf,
} from '../package/content-control-nodes.ts';
import { findNode, parentNodeOf } from '../package/ooxml-edit.ts';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { contentControlContentOf } from './tree-op-nodes.ts';
import type { TreeOpRejection } from './tree-op-types.ts';

/**
 * Where a new block goes. @public
 *
 * `inside` names a block-level control showing its prompt: what is written replaces the prompt
 * whole, as a node written into an inline prompt does. A control that holds content has places
 * of its own — before and after each of its blocks — so `inside` one is refused.
 */
export type BlockPlace =
  | { readonly before: string }
  | { readonly after: string }
  | { readonly inside: string };

/** Where the new blocks land: at `index` among `holder`'s children. */
export interface BlockLanding {
  readonly holder: OoxmlElement;
  readonly index: number;
  /** The prompt-showing control the place is inside, whose prompt the new blocks replace. */
  readonly prompt: OoxmlElement | null;
}

/** The node a well-formed place names, or null for a place of no known shape. */
export function blockPlaceAnchorOf(place: BlockPlace | null | undefined): string | null {
  if (typeof place !== 'object' || place === null) return null;
  const entries = Object.entries(place);
  if (entries.length !== 1) return null;
  const [key, id] = entries[0]!;
  if (key !== 'before' && key !== 'after' && key !== 'inside') return null;
  return typeof id === 'string' && id.length > 0 ? id : null;
}

/** Whether the place is inside a control, as opposed to beside a block. */
export function isInsidePlace(place: BlockPlace): place is { readonly inside: string } {
  return 'inside' in place;
}

/** A block of a story: a paragraph, a table, or a control that holds blocks. */
export function isBlockNode(node: OoxmlNode): node is OoxmlElement {
  if (node.kind === 'paragraph' || node.kind === 'table') return true;
  return node.kind === 'contentControl' && contentControlLevelOf(node) === 'block';
}

/** The block an id names, or why it names none. */
export function blockOf(part: OoxmlPart, id: string): OoxmlElement | TreeOpRejection {
  const node = typeof id === 'string' && id.length > 0 ? findNode(part, id) : null;
  if (!node) return 'unknown-block';
  return isBlockNode(node) ? node : 'not-a-block';
}

/**
 * A container whose content is a control's prompt has no place beside it: the prompt is
 * replaced whole, through `inside`.
 */
export function holdsPrompt(part: OoxmlPart, holder: OoxmlNode): boolean {
  if (holder.kind !== 'contentControlContent') return false;
  const owner = parentNodeOf(part, holder.id);
  return owner?.kind === 'contentControl' && contentControlPropertiesOf(owner).showingPlaceholder;
}

/** Resolve a place for the validator and the applier alike, so `can` and the write agree. */
export function blockLandingOf(part: OoxmlPart, place: BlockPlace): BlockLanding | TreeOpRejection {
  if (blockPlaceAnchorOf(place) === null) return 'invalidArgs';
  if (isInsidePlace(place)) return insideLanding(part, place.inside);
  const anchorId = 'before' in place ? place.before : place.after;
  const anchor = blockOf(part, anchorId);
  if (typeof anchor === 'string') return anchor;
  const holder = parentNodeOf(part, anchor.id);
  if (!holder) return 'not-a-block';
  if (holdsPrompt(part, holder)) return 'invalidArgs';
  const at = holder.children.findIndex((child) => child.id === anchor.id);
  return { holder, index: 'before' in place ? at : at + 1, prompt: null };
}

function insideLanding(part: OoxmlPart, controlId: string): BlockLanding | TreeOpRejection {
  const control = findNode(part, controlId);
  if (!control) return 'unknown-content-control';
  if (control.kind !== 'contentControl') return 'not-a-content-control';
  if (contentControlLevelOf(control) !== 'block') return 'invalidArgs';
  if (!contentControlPropertiesOf(control).showingPlaceholder) return 'invalidArgs';
  const content = contentControlContentOf(control);
  if (!content) return 'unsupported';
  return { holder: content, index: 0, prompt: control };
}

/**
 * Whether a cell would still end with a paragraph (§17.4.66) once a control takes the place at
 * `index`: some block must stand after it. A cell whose last block were a control is markup Word
 * rejects, and a cell the caret could not leave.
 */
export function cellKeepsItsLastBlock(holder: OoxmlElement, index: number): boolean {
  return holder.kind !== 'tableCell' || holder.children.slice(index).some(isBlockNode);
}
