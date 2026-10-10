// Authoring a NEW content control (store lane).
//
// Two shapes, one op. A range becomes a WRAPPER around the characters that are already
// there; a caret becomes an EMPTY control showing its type's prompt, which is what Word's
// Developer tab inserts when nothing is selected. Both are one transaction and one undo
// step, and both leave every other node in the paragraph where it was. Either lands in the
// paragraph, or in the place the caller names — inside an inline control, or beside one.
//
// Split out of `tree-op-content-controls.ts` so the insertion paths have room to be read
// side by side rather than to fit under a line cap.

import {
  allocateContentControlId,
  lockForbidsEdit,
  orderedContentControlProperties,
} from '../package/content-control-nodes.ts';
import {
  createNodeIdAllocator,
  findNode,
  replaceChildren,
  type EditOptions,
} from '../package/ooxml-edit.ts';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { splitRunsAt } from './tree-op-apply.ts';
import {
  cloneWithFreshIds,
  contentControlEffect,
  contentControlLockAt,
  editedProperties,
  promptFor,
  textRun,
  wmlElement,
  type InsertableContentControlKind,
} from './tree-op-content-controls.ts';
import { contentControlContentOf, fromEdit, runPropertiesNodeOf } from './tree-op-nodes.ts';
import {
  atomSpanLookup,
  inlineDestinationOf,
  inlineDestinationRefusal,
  inlineLandingAt,
  type InlineDestination,
} from './tree-op-inline-destination.ts';
import {
  indivisibleAt,
  paragraphLength,
  paragraphOffsetIndex,
  splitsSurrogate,
  type ParagraphOffsetIndex,
} from './tree-op-segments.ts';
import type { TreeDocOp, TreeOpResult } from './tree-op-types.ts';

/** The ECMA-376 `CT_SdtPr` type element each insertable kind writes. */
const TYPE_ELEMENT_FOR: Readonly<Record<InsertableContentControlKind, string>> = {
  richText: 'richText',
  plainText: 'text',
  dropDownList: 'dropDownList',
  comboBox: 'comboBox',
  date: 'date',
};

type InsertOp = Extract<TreeDocOp, { op: 'insertContentControl' }>;

/** What a new control's `w:sdtPr` says, at either level. */
export interface NewControlProperties extends Pick<InsertOp, 'tag' | 'alias' | 'lock' | 'type'> {
  /**
   * A wrapper holds content the caller chose, so it shows no prompt; an empty control holds
   * nothing but one, and the flag is what makes the first keystroke replace it whole rather
   * than append to it.
   */
  readonly showingPlaceholder: boolean;
}

/** `w:sdtPr` + the type element, in schema order, with the caller's metadata on it. */
export function propertiesFor(
  part: OoxmlPart,
  control: NewControlProperties,
  nextId: () => string
): OoxmlElement {
  // Inside the store transaction, so the collaboration actor is already bound. A second
  // local max+1 here would ignore that bind and collide the moment two peers insert.
  const allocated = allocateContentControlId(part.root);
  const properties = editedProperties(
    undefined,
    {
      ...(control.tag === undefined ? {} : { tag: control.tag }),
      ...(control.alias === undefined ? {} : { alias: control.alias }),
      ...(allocated === null ? {} : { id: allocated }),
      ...(control.lock === undefined ? {} : { lock: control.lock }),
      ...(control.showingPlaceholder ? { showingPlaceholder: true } : {}),
    },
    nextId
  );
  const typed = wmlElement(nextId, TYPE_ELEMENT_FOR[control.type]);
  return {
    ...properties,
    children: orderedContentControlProperties([...properties.children, typed]),
  } as OoxmlElement;
}

export function controlElement(
  properties: OoxmlElement,
  content: readonly OoxmlNode[],
  nextId: () => string
): OoxmlElement {
  const wrapped = wmlElement(nextId, 'sdtContent', {
    kind: 'contentControlContent' as OoxmlNode['kind'],
    children: content,
  });
  return wmlElement(nextId, 'sdt', {
    kind: 'contentControl' as OoxmlNode['kind'],
    children: [properties, wrapped],
  });
}

export function applyInsertContentControl(
  part: OoxmlPart,
  op: InsertOp,
  options?: EditOptions
): TreeOpResult {
  const paragraph = findNode(part, op.paragraphId);
  if (!paragraph) return { ok: false, reason: 'unknown-paragraph' };
  if (paragraph.kind !== 'paragraph') return { ok: false, reason: 'not-a-paragraph' };
  if (lockForbidsEdit(contentControlLockAt(part, op.paragraphId))) {
    return { ok: false, reason: 'locked' };
  }

  // `start === end` is a CARET, not an empty range: it authors the prompt-showing control
  // Word's Developer tab inserts with nothing selected. Refusing it here is what forced a
  // host to select a character first and then delete it, which is two undo steps and a
  // document edit nobody asked for.
  if (op.start < 0 || op.end > paragraphLength(paragraph) || op.start > op.end) {
    return { ok: false, reason: 'invalid-range' };
  }
  if (splitsSurrogate(paragraph, op.start) || splitsSurrogate(paragraph, op.end)) {
    return { ok: false, reason: 'splits-surrogate-pair' };
  }
  // `validateTreeOp` answers these too, which is what lets `can` predict the refusal. Repeated
  // here because the applier is reachable on its own and must fail closed rather than emit the
  // control beside the container the caller pointed into.
  const destination = inlineDestinationOf(op);
  const refused =
    destination === null
      ? indivisibleAt(paragraph, op.start) || indivisibleAt(paragraph, op.end)
        ? 'indivisible-content'
        : null
      : inlineDestinationRefusal(part, op.paragraphId, op, op);
  if (refused) return { ok: false, reason: refused };
  return op.start === op.end
    ? insertEmptyContentControl(part, op, destination, options)
    : wrapRangeInContentControl(part, op, destination, options);
}

/**
 * Wrap `[start, end)` in a new control.
 *
 * A control is a SIBLING of runs, never a thing inside one, so a range ending mid-run splits
 * that run at both edges first. The characters and their formatting are the ones that were
 * there; only the run boundaries move. The wrapped children are the paragraph's, or the named
 * owner's when the range lies inside an inline control.
 */
function wrapRangeInContentControl(
  part: OoxmlPart,
  op: InsertOp,
  destination: InlineDestination | null,
  options?: EditOptions
): TreeOpResult {
  let current = part;
  for (const edge of [op.end, op.start]) {
    const target = findNode(current, op.paragraphId);
    if (!target || target.kind !== 'paragraph') return { ok: false, reason: 'tree-invariant' };
    const split = splitRunsAt(current, target, edge, options);
    if (!split.ok) return { ok: false, reason: split.reason };
    current = split.part;
  }

  const reloaded = findNode(current, op.paragraphId);
  if (!reloaded || reloaded.kind !== 'paragraph') return { ok: false, reason: 'tree-invariant' };
  const holder =
    destination === null
      ? reloaded
      : contentControlContentOf(findNode(current, destination.controlId)!);
  if (!holder) return { ok: false, reason: 'unsupported' };
  const index = paragraphOffsetIndex(reloaded);
  const wrapped: OoxmlNode[] = [];
  let covered = false;
  // A field's chrome sits at ZERO length at the field's offset and belongs to its atom, as on
  // the caret path below: skipping it left the instruction and the end marker outside the
  // control, cutting the field in two.
  const atomSpanOf = atomSpanLookup(index);
  for (const child of holder.children) {
    const own = index.spanOf(child);
    const span = own && own.start !== own.end ? own : (atomSpanOf(child) ?? own);
    if (!span || span.start === span.end) continue;
    if (span.start >= op.start && span.end <= op.end) {
      wrapped.push(child);
      covered = true;
      continue;
    }
    if (span.start < op.end && span.end > op.start) return { ok: false, reason: 'invalid-range' };
  }
  if (!covered) return { ok: false, reason: 'invalid-range' };

  const nextId = createNodeIdAllocator(current);
  const control = controlElement(
    propertiesFor(current, { ...op, showingPlaceholder: false }, nextId),
    wrapped,
    nextId
  );

  const wrappedIds = new Set(wrapped.map((child) => child.id));
  let placed = false;
  const children: OoxmlNode[] = [];
  for (const child of holder.children) {
    if (!wrappedIds.has(child.id)) {
      children.push(child);
      continue;
    }
    if (!placed) {
      children.push(control);
      placed = true;
    }
  }
  return fromEdit(
    replaceChildren(current, holder.id, children, options),
    contentControlEffect(reloaded.id, 'flow-structural')
  );
}

/**
 * The `w:rPr` the prompt run inherits: the run on the LEFT of the caret, falling back to the
 * one on the right at paragraph start.
 *
 * The same rule the inline custom-node insert follows, for the same reason — a field dropped
 * into a heading must not come out body-sized.
 */
export function inheritedRunProperties(
  part: OoxmlPart,
  index: ParagraphOffsetIndex,
  offset: number,
  nextId: () => string
): OoxmlNode | undefined {
  const segments = index.segments;
  let anchor = null as (typeof segments)[number] | null;
  for (const segment of segments) {
    if (segment.end <= offset && segment.end > 0) anchor = segment;
  }
  anchor ??= segments.find((segment) => segment.start >= offset) ?? null;
  const run = anchor ? findNode(part, anchor.runId) : null;
  const properties = run && run.kind === 'run' ? runPropertiesNodeOf(run) : null;
  return properties ? cloneWithFreshIds(properties, nextId) : undefined;
}

/**
 * Insert an EMPTY control at a caret, holding its type's prompt.
 *
 * Word's own gesture: the control arrives showing "Click here to enter text." with
 * `w:showingPlcHdr` set, so the first character typed replaces the prompt whole. That
 * transition already exists in the applier for `insertText`; this is the other end of it.
 * The caret's place is the paragraph's, or the one the caller named (`inside`, `beside`).
 */
function insertEmptyContentControl(
  part: OoxmlPart,
  op: InsertOp,
  destination: InlineDestination | null,
  options?: EditOptions
): TreeOpResult {
  const landed = inlineLandingAt(part, op.paragraphId, op.start, destination, options);
  if (!landed.ok) return landed;
  const { landing } = landed;
  const holder = findNode(landing.part, landing.holderId);
  if (!holder || holder.kind === 'textValue') return { ok: false, reason: 'tree-invariant' };

  const nextId = createNodeIdAllocator(landing.part);
  const prompt = textRun(
    nextId,
    promptFor(op.type, options),
    inheritedRunProperties(
      landing.part,
      paragraphOffsetIndex(landing.paragraph),
      landing.offset,
      nextId
    )
  );
  const control = controlElement(
    propertiesFor(landing.part, { ...op, showingPlaceholder: true }, nextId),
    [prompt],
    nextId
  );
  const children = [
    ...holder.children.slice(0, landing.index),
    control,
    ...holder.children.slice(landing.index),
  ];
  return fromEdit(
    replaceChildren(landing.part, holder.id, children, options),
    contentControlEffect(landing.paragraph.id, 'flow-structural')
  );
}
