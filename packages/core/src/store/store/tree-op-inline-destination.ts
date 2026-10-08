// Where a new control or an inline fragment lands when the caller NAMES the place.
//
// An offset at an edge several inline controls share is several places, and an offset strictly
// inside an inline control has no paragraph-level sibling position at all. `insertText` already
// takes the two names that say which: `inside` a control (its own content) and `beside` one (its
// sibling at that edge). Authoring a control or landing a fragment takes the same two, resolved
// here once for the validation and the applier — two copies of "where" is how a refusal ends up
// reasoning about a different place than the write.

import type { EditOptions } from '../package/ooxml-edit.ts';
import { findNode } from '../package/ooxml-edit.ts';
import type {
  OoxmlElement,
  OoxmlNode,
  OoxmlParagraphNode,
  OoxmlPart,
} from '../package/ooxml-tree.ts';
import { contentControlPropertiesOf } from '../package/content-control-nodes.ts';
import { splitRunsAt } from './tree-op-apply.ts';
import { siteBesideControl } from './tree-op-beside.ts';
import { clearPlaceholder, placeholderControlForInsertion } from './tree-op-content-controls.ts';
import {
  contentControlContentOf,
  findContentControl,
  isParagraphPropertiesNode,
} from './tree-op-nodes.ts';
import {
  isParagraph,
  paragraphOffsetIndex,
  type OffsetSpan,
  type ParagraphOffsetIndex,
} from './tree-op-segments.ts';
import { namedOwnerRefusal } from './tree-op-validate.ts';
import { contentControlAtCaret, holds, rejectContentEdit } from './tree-op-validate-controls.ts';
import type { TreeOpRejection } from './tree-op-types.ts';

/** The place a caller names: a control's own content, or the sibling slot at one of its edges. */
export type InlineDestination =
  | { readonly kind: 'inside'; readonly controlId: string }
  | {
      readonly kind: 'beside';
      readonly controlId: string;
      readonly side: 'before' | 'after';
    };

/** The two optional fields an op carries, as `insertText` spells them. */
export interface InlineDestinationFields {
  /** Land in this inline control's own content, at the offset. */
  readonly inside?: string;
  /** Land as this inline control's sibling, at its `side` edge, which must be the offset. */
  readonly beside?: { readonly controlId: string; readonly side: 'before' | 'after' };
}

/** Where the new nodes go: before the child at `index` of `holderId`. */
export interface InlineLanding {
  readonly part: OoxmlPart;
  readonly paragraph: OoxmlParagraphNode;
  readonly holderId: string;
  readonly index: number;
  /** The offset the nodes land at, which moves when an emptied prompt is the destination. */
  readonly offset: number;
}

export type InlineLandingResult =
  | { readonly ok: true; readonly landing: InlineLanding }
  | { readonly ok: false; readonly reason: TreeOpRejection };

export function inlineDestinationOf(fields: InlineDestinationFields): InlineDestination | null {
  if (fields.beside !== undefined) return { kind: 'beside', ...fields.beside };
  return fields.inside === undefined ? null : { kind: 'inside', controlId: fields.inside };
}

/**
 * What a named destination refuses, for `[start, end)`: the same answers the applier gives, so
 * `can` never calls a write live that then fails. `beside` names a caret; `inside` names a range
 * within the control's own offsets whose edges do not cut what the control holds.
 */
export function inlineDestinationRefusal(
  part: OoxmlPart,
  paragraphId: string,
  range: { readonly start: number; readonly end: number },
  fields: InlineDestinationFields
): TreeOpRejection | null {
  if (fields.inside !== undefined && fields.beside !== undefined) return 'invalidArgs';
  const destination = inlineDestinationOf(fields);
  const paragraph = findNode(part, paragraphId);
  if (destination === null || !paragraph || !isParagraph(paragraph)) return null;
  const refused =
    destination.kind === 'beside'
      ? besideRefusal(paragraph, range, destination)
      : insideRefusal(part, paragraph, range, destination.controlId);
  return refused ?? rejectContentEdit(part, paragraph, range.start, range.end);
}

function besideRefusal(
  paragraph: OoxmlParagraphNode,
  range: { readonly start: number; readonly end: number },
  beside: Extract<InlineDestination, { kind: 'beside' }>
): TreeOpRejection | null {
  if (range.start !== range.end) return 'invalidArgs';
  if (beside.side !== 'before' && beside.side !== 'after') return 'invalidArgs';
  return siteBesideControl(paragraph, range.start, beside) ? null : 'unknown-content-control';
}

function insideRefusal(
  part: OoxmlPart,
  paragraph: OoxmlParagraphNode,
  range: { readonly start: number; readonly end: number },
  controlId: string
): TreeOpRejection | null {
  const named =
    namedOwnerRefusal(part, paragraph.id, range.start, controlId) ??
    namedOwnerRefusal(part, paragraph.id, range.end, controlId);
  if (named) return named;
  // A block control that holds the paragraph is not a place inside the paragraph: the paragraph
  // itself is, and that is what no destination already means.
  if (!holds(paragraph, controlId)) return 'invalidArgs';
  const owner = findNode(part, controlId)!;
  if (contentControlPropertiesOf(owner).showingPlaceholder && range.start !== range.end) {
    return 'invalid-range';
  }
  const content = contentControlContentOf(owner);
  if (!content) return 'unsupported';
  const index = paragraphOffsetIndex(paragraph);
  return cutsHeldContent(index, content, range.start) || cutsHeldContent(index, content, range.end)
    ? 'indivisible-content'
    : null;
}

/**
 * Whether `offset` falls strictly inside something a holder keeps whole: a child that is not a
 * run, or an atom (a field) anywhere inside it. A run divides at any offset.
 */
function cutsHeldContent(
  index: ParagraphOffsetIndex,
  holder: OoxmlElement,
  offset: number
): boolean {
  const strictlyInside = (span: OffsetSpan | null | undefined): boolean =>
    !!span && span.start < offset && offset < span.end;
  const atoms = index.segments.filter(
    (segment) => segment.removeNodeIds && holds(holder, segment.node.id)
  );
  return (
    atoms.some(strictlyInside) ||
    holder.children.some((child) => child.kind !== 'run' && strictlyInside(index.spanOf(child)))
  );
}

/**
 * Resolve the landing of a caret for the applier. Without a destination it is the paragraph
 * itself; `beside` is the sibling slot at the control's edge; `inside` is the control's content,
 * after its prompt is emptied when it shows one — the first thing written into a prompt
 * replaces it, as the first keystroke does.
 */
export function inlineLandingAt(
  part: OoxmlPart,
  paragraphId: string,
  offset: number,
  destination: InlineDestination | null,
  options?: EditOptions
): InlineLandingResult {
  if (destination?.kind === 'beside') return besideLanding(part, paragraphId, offset, destination);
  const prepared =
    destination === null
      ? { part, offset }
      : emptiedOwner(part, paragraphId, offset, destination.controlId, options);
  if (prepared === null) return { ok: false, reason: 'unsupported' };
  const split = splitAt(prepared.part, paragraphId, prepared.offset, options);
  if (!split.ok) return split;
  const holder =
    destination === null
      ? split.paragraph
      : contentControlContentOf(findNode(split.part, destination.controlId)!);
  if (!holder) return { ok: false, reason: 'unsupported' };
  return {
    ok: true,
    landing: {
      part: split.part,
      paragraph: split.paragraph,
      holderId: holder.id,
      index: childrenBefore(
        paragraphOffsetIndex(split.paragraph),
        holder.children,
        prepared.offset
      ),
      offset: prepared.offset,
    },
  };
}

function besideLanding(
  part: OoxmlPart,
  paragraphId: string,
  offset: number,
  beside: Extract<InlineDestination, { kind: 'beside' }>
): InlineLandingResult {
  const paragraph = findNode(part, paragraphId);
  if (!paragraph || !isParagraph(paragraph)) return { ok: false, reason: 'unknown-paragraph' };
  const site = siteBesideControl(paragraph, offset, beside);
  if (site?.kind !== 'newRun' || site.index === undefined) {
    return { ok: false, reason: 'unknown-content-control' };
  }
  return {
    ok: true,
    landing: { part, paragraph, holderId: site.holder.id, index: site.index, offset },
  };
}

/**
 * The prompt an insertion types over, and the offset its text then starts at.
 *
 * Decided by where the text lands, as the first keystroke into a prompt replaces it: the named
 * control's own prompt `inside`, none `beside` a control, and the offset's own rule only when the
 * caller names no place. The prompt of a neighbour sharing the offset is never the one replaced.
 */
export function promptTypedOver(
  part: OoxmlPart,
  paragraphId: string,
  offset: number,
  fields: InlineDestinationFields
): { readonly control: OoxmlNode; readonly offset: number } | null {
  if (fields.beside !== undefined) return null;
  if (fields.inside === undefined) return placeholderControlForInsertion(part, paragraphId, offset);
  const owner = findNode(part, fields.inside);
  const paragraph = findNode(part, paragraphId);
  if (!owner || !paragraph || !isParagraph(paragraph)) return null;
  if (!contentControlPropertiesOf(owner).showingPlaceholder) return null;
  const start = holds(owner, paragraphId)
    ? 0
    : paragraphOffsetIndex(paragraph).spanOf(owner)?.start;
  return start === undefined ? null : { control: owner, offset: start };
}

/** The control an insertion at a caret edits: the one it names, none beside one, or the offset's. */
export function insertionOwnerAt(
  part: OoxmlPart,
  paragraph: OoxmlParagraphNode,
  offset: number,
  fields: InlineDestinationFields,
  bias?: 'left' | 'right'
): OoxmlElement | null {
  if (fields.beside !== undefined) return null;
  if (fields.inside !== undefined) return findContentControl(part, fields.inside);
  return contentControlAtCaret(part, paragraph, offset, offset, bias);
}

/** The owner with its prompt emptied, and the offset its content now starts at. */
function emptiedOwner(
  part: OoxmlPart,
  paragraphId: string,
  offset: number,
  controlId: string,
  options?: EditOptions
): { readonly part: OoxmlPart; readonly offset: number } | null {
  if (!findNode(part, controlId)) return null;
  const prompt = promptTypedOver(part, paragraphId, offset, { inside: controlId });
  if (prompt === null) return { part, offset };
  const emptied = clearPlaceholder(part, controlId, options, paragraphId);
  return emptied === null ? null : { part: emptied, offset: prompt.offset };
}

function splitAt(
  part: OoxmlPart,
  paragraphId: string,
  offset: number,
  options?: EditOptions
):
  | { readonly ok: true; readonly part: OoxmlPart; readonly paragraph: OoxmlParagraphNode }
  | { readonly ok: false; readonly reason: TreeOpRejection } {
  const paragraph = findNode(part, paragraphId);
  if (!paragraph || !isParagraph(paragraph)) return { ok: false, reason: 'unknown-paragraph' };
  const split = splitRunsAt(part, paragraph, offset, options);
  if (!split.ok) return split;
  const reloaded = findNode(split.part, paragraphId);
  if (!reloaded || !isParagraph(reloaded)) return { ok: false, reason: 'tree-invariant' };
  return { ok: true, part: split.part, paragraph: reloaded };
}

/**
 * How many of a holder's children sit before `offset`, once the run there is split.
 *
 * A field's chrome — instruction, separator, end marker — sits at ZERO length at the field's
 * offset and belongs to its atom, so the atom's span answers for each of those nodes; a truly
 * zero-length node (a bookmark) takes the side its position puts it on. Nothing straddles the
 * offset: the run was split, and a cut through anything else was refused.
 */
export function childrenBefore(
  index: ParagraphOffsetIndex,
  children: readonly OoxmlNode[],
  offset: number
): number {
  const atomSpanOf = atomSpanLookup(index);
  let before = 0;
  let cursor = 0;
  for (const child of children) {
    if (isParagraphPropertiesNode(child)) {
      before += 1;
      continue;
    }
    const own = index.spanOf(child);
    const span = own && own.start !== own.end ? own : (atomSpanOf(child) ?? own);
    if (!span || span.start === span.end) {
      if (cursor < offset) before += 1;
      continue;
    }
    cursor = span.end;
    if (span.end > offset) break;
    before += 1;
  }
  return before;
}

/**
 * The span of the ATOM a child belongs to, for the chrome an atom is spelt with.
 *
 * `removeNodeIds` names the elements — `w:fldChar`, `w:instrText` — rather than the runs
 * holding them, so a child is matched by itself or by what it holds. One level is enough:
 * field chrome is a run wrapping exactly one of those elements.
 */
export function atomSpanLookup(
  index: ParagraphOffsetIndex
): (child: OoxmlNode) => OffsetSpan | null {
  const byNodeId = new Map<string, OffsetSpan>();
  for (const segment of index.segments) {
    if (!segment.removeNodeIds) continue;
    const span = { start: segment.start, end: segment.end };
    for (const id of segment.removeNodeIds) byNodeId.set(id, span);
  }
  if (byNodeId.size === 0) return () => null;
  return (child) => {
    const own = byNodeId.get(child.id);
    if (own) return own;
    if (child.kind === 'textValue') return null;
    for (const inner of child.children) {
      const found = byNodeId.get(inner.id);
      if (found) return found;
    }
    return null;
  };
}
