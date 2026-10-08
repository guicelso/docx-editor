// Merge fields a host authors: `MERGEFIELD "<name>"`, the format's "place that receives a value".
//
// The NAME is the host's — this engine never reads meaning into it — and the instruction around
// it is the engine's: one spelling, quoted, with `\* MERGEFORMAT` so a Word that refreshes the
// field keeps the result's formatting. A host that wrote the instruction string itself had to
// know the field-code quoting rules and build five runs by hand; this module is the one place
// that does both.
//
// The field is five runs (`begin`, instruction, `separate`, the cached result, `end`), each
// carrying the formatting in force at the caret, so the result reads like the text around it
// and a reader that walks fields run by run finds every piece in its own `w:r`. It lands where
// the caret's place says — the paragraph, or the place a caller names (`inside`, `beside`) —
// through the same landing every inline insertion uses, so no paragraph is split and joined and
// no control beside the caret is touched.

import { fieldOnOffAttribute, isFldCharNode, isInstrText } from '../package/field-nodes.ts';
import { WML_NAMESPACE_URI } from '../package/ooxml-shared.ts';
import {
  createNodeIdAllocator,
  findNode,
  insertChildren,
  replaceNode,
  type EditOptions,
} from '../package/ooxml-edit.ts';
import type {
  OoxmlElement,
  OoxmlNode,
  OoxmlParagraphNode,
  OoxmlPart,
} from '../package/ooxml-tree.ts';
import { isValidXmlText } from '../package/sinks.ts';
import { inheritedRunProperties } from './tree-op-content-control-insert.ts';
import { cloneWithFreshIds, textRun } from './tree-op-content-controls.ts';
import {
  applyFieldResults,
  fieldResultUpdateRefusal,
  locateFieldResults,
  MAX_FIELD_RESULT_TEXT_CHARS,
} from './tree-op-field-results.ts';
import { fldChar, instrText } from './tree-op-fields.ts';
import {
  inlineDestinationOf,
  inlineDestinationRefusal,
  inlineLandingAt,
  promptTypedOver,
  type InlineDestination,
} from './tree-op-inline-destination.ts';
import { fromEdit, TEXT_DEPS } from './tree-op-nodes.ts';
import {
  indivisibleAt,
  isParagraph,
  paragraphLength,
  paragraphOffsetIndex,
  splitsSurrogate,
} from './tree-op-segments.ts';
import type { TreeOpRejection, TreeOpResult } from './tree-op-types.ts';
import { rejectContentEdit } from './tree-op-validate-controls.ts';

/** Author a merge field at a caret. */
export interface InsertMergeFieldOp {
  readonly op: 'insertMergeField';
  readonly paragraphId: string;
  readonly offset: number;
  /** The field's name, written quoted into the instruction. The host's vocabulary, never parsed. */
  readonly name: string;
  /** The cached result the page shows until the field is filled. */
  readonly result: string;
  /** Land in this inline control's own content, at the offset — the name `insertText` takes. */
  readonly inside?: string;
  /** Land as this inline control's sibling at its edge, which must be the offset. */
  readonly beside?: { readonly controlId: string; readonly side: 'before' | 'after' };
}

/** Rewrite a merge field's name and cached result in place: its runs and their formatting stay. */
export interface SetMergeFieldOp {
  readonly op: 'setMergeField';
  readonly paragraphId: string;
  /** The field's begin `w:fldChar`, the identity every field read reports. */
  readonly fieldNodeId: string;
  readonly name: string;
  readonly result: string;
}

export type MergeFieldOp = InsertMergeFieldOp | SetMergeFieldOp;

/** A merge field of a paragraph: its identity, its name and the offsets its atom occupies. */
export interface MergeFieldRange {
  readonly fieldNodeId: string;
  readonly name: string;
  readonly start: number;
  readonly end: number;
}

/** A name is one field-code argument: bounded, printable, and free of the quote and the escape. */
export const MAX_MERGE_FIELD_NAME_CHARS = 2048;

export function isMergeFieldName(name: unknown): name is string {
  return (
    typeof name === 'string' &&
    name.length > 0 &&
    name.length <= MAX_MERGE_FIELD_NAME_CHARS &&
    isValidXmlText(name) &&
    !/["\\\t\r\n]/u.test(name)
  );
}

/** The instruction, without the padding space `instrText` adds on each side. */
function mergeFieldCode(name: string): string {
  return `MERGEFIELD "${name}" \\* MERGEFORMAT`;
}

/**
 * The name a merge field's instruction carries, or `null` when the instruction is not one.
 *
 * Quoted or bare, as Word writes either; switches after it are not part of the name.
 */
export function mergeFieldNameOf(instruction: string): string | null {
  const quoted = /^\s*MERGEFIELD\s+"([^"]*)"/u.exec(instruction);
  if (quoted) return quoted[1]!.length > 0 ? quoted[1]! : null;
  const bare = /^\s*MERGEFIELD\s+([^\s"\\]+)/u.exec(instruction);
  return bare ? bare[1]! : null;
}

/** The merge fields of one paragraph in document order, each with its atom's offsets. */
export function mergeFieldsOf(paragraph: OoxmlParagraphNode): readonly MergeFieldRange[] {
  const atoms = new Map(
    paragraphOffsetIndex(paragraph)
      .segments.filter((segment) => segment.removeNodeIds !== undefined)
      .map((segment) => [segment.node.id, segment] as const)
  );
  return locateFieldResults(paragraph).flatMap((field) => {
    const name = mergeFieldNameOf(field.instruction);
    const atom = atoms.get(field.fieldNodeId);
    return name === null || atom === undefined
      ? []
      : [{ fieldNodeId: field.fieldNodeId, name, start: atom.start, end: atom.end }];
  });
}

function validResult(result: unknown): result is string {
  return (
    typeof result === 'string' &&
    result.length <= MAX_FIELD_RESULT_TEXT_CHARS &&
    isValidXmlText(result) &&
    !/[\t\r\n]/u.test(result)
  );
}

export function isMergeFieldOp(op: { readonly op: string }): op is MergeFieldOp {
  return op.op === 'insertMergeField' || op.op === 'setMergeField';
}

export function validateMergeFieldOp(part: OoxmlPart, op: MergeFieldOp): TreeOpRejection | null {
  if (!isMergeFieldName(op.name) || !validResult(op.result)) return 'invalidArgs';
  const paragraph = findNode(part, op.paragraphId);
  if (!paragraph) return 'unknown-paragraph';
  if (!isParagraph(paragraph)) return 'not-a-paragraph';
  return op.op === 'insertMergeField'
    ? insertRefusal(part, paragraph, op)
    : rewriteRefusal(part, paragraph, op);
}

function insertRefusal(
  part: OoxmlPart,
  paragraph: OoxmlParagraphNode,
  op: InsertMergeFieldOp
): TreeOpRejection | null {
  if (!Number.isInteger(op.offset) || op.offset < 0 || op.offset > paragraphLength(paragraph)) {
    return 'offset-out-of-range';
  }
  if (splitsSurrogate(paragraph, op.offset)) return 'splits-surrogate-pair';
  const caret = { start: op.offset, end: op.offset };
  if (op.inside !== undefined || op.beside !== undefined) {
    return inlineDestinationRefusal(part, op.paragraphId, caret, op);
  }
  if (indivisibleAt(paragraph, op.offset)) return 'indivisible-content';
  return rejectContentEdit(part, paragraph, op.offset, op.offset);
}

function rewriteRefusal(
  part: OoxmlPart,
  paragraph: OoxmlParagraphNode,
  op: SetMergeFieldOp
): TreeOpRejection | null {
  const field = mergeFieldsOf(paragraph).find((each) => each.fieldNodeId === op.fieldNodeId);
  const located = locateFieldResults(paragraph).find((each) => each.fieldNodeId === op.fieldNodeId);
  const anchor = findNode(part, op.fieldNodeId);
  if (!field || !located?.rewritable || !anchor) return 'invalidArgs';
  if (fieldOnOffAttribute(anchor, 'fldLock') === true) return 'locked';
  return (
    fieldResultUpdateRefusal(part, op.paragraphId) ??
    rejectContentEdit(part, paragraph, field.start, field.end)
  );
}

export function applyMergeFieldOp(
  part: OoxmlPart,
  op: MergeFieldOp,
  options?: EditOptions
): TreeOpResult {
  const refused = validateMergeFieldOp(part, op);
  if (refused) return { ok: false, reason: refused };
  return op.op === 'insertMergeField'
    ? applyInsertMergeField(part, op, options)
    : applySetMergeField(part, op, options);
}

/**
 * A caret in a control showing its prompt writes into that control, replacing the prompt — the
 * rule the first keystroke follows, decided by the same function `insertText` uses.
 */
function destinationOf(part: OoxmlPart, op: InsertMergeFieldOp): InlineDestination | null {
  const named = inlineDestinationOf(op);
  if (named !== null) return named;
  const prompt = promptTypedOver(part, op.paragraphId, op.offset, op);
  return prompt === null ? null : { kind: 'inside', controlId: prompt.control.id };
}

function applyInsertMergeField(
  part: OoxmlPart,
  op: InsertMergeFieldOp,
  options?: EditOptions
): TreeOpResult {
  const landed = inlineLandingAt(part, op.paragraphId, op.offset, destinationOf(part, op), options);
  if (!landed.ok) return landed;
  const { landing } = landed;
  const nextId = createNodeIdAllocator(landing.part);
  const formatting = inheritedRunProperties(
    landing.part,
    paragraphOffsetIndex(landing.paragraph),
    landing.offset,
    nextId
  );
  const runOf = (child: OoxmlNode): OoxmlNode => fieldRun(nextId, child, formatting);
  const runs: readonly OoxmlNode[] = [
    runOf(fldChar(nextId, 'begin')),
    runOf(instrText(nextId, mergeFieldCode(op.name))),
    runOf(fldChar(nextId, 'separate')),
    textRun(nextId, op.result, formatting && cloneWithFreshIds(formatting, nextId)),
    runOf(fldChar(nextId, 'end')),
  ];
  return fromEdit(insertChildren(landing.part, landing.holderId, landing.index, runs, options), {
    dirty: [landing.paragraph.id],
    created: [],
    deleted: [],
    dependencyKeys: TEXT_DEPS,
    impact: 'text-local',
  });
}

/** One run holding one field piece, with its own copy of the caret's formatting. */
function fieldRun(
  nextId: () => string,
  child: OoxmlNode,
  formatting: OoxmlNode | undefined
): OoxmlNode {
  return {
    id: nextId(),
    kind: 'run',
    namespaceUri: WML_NAMESPACE_URI,
    localName: 'r',
    prefix: 'w',
    namespaceBindings: [],
    attributes: [],
    children: formatting ? [cloneWithFreshIds(formatting, nextId), child] : [child],
  } as unknown as OoxmlNode;
}

/** The instruction's text is rewritten on its first `w:instrText`; the others are emptied. */
function applySetMergeField(
  part: OoxmlPart,
  op: SetMergeFieldOp,
  options?: EditOptions
): TreeOpResult {
  let current = part;
  let wrote = false;
  for (const node of instructionNodesOf(part, op.paragraphId, op.fieldNodeId)) {
    const value = wrote ? '' : ` ${mergeFieldCode(op.name)} `;
    wrote = true;
    const updated = {
      ...node,
      children: node.children.map((child) =>
        child.kind === 'textValue' ? { ...child, value } : child
      ),
    } as OoxmlNode;
    const replaced = replaceNode(current, node.id, updated, options);
    if (!replaced.ok) return { ok: false, reason: 'tree-invariant' };
    current = replaced.part;
  }
  if (!wrote) return { ok: false, reason: 'invalidArgs' };
  const refreshed = applyFieldResults(
    current,
    {
      op: 'refreshFieldResults',
      updates: [{ paragraphId: op.paragraphId, fieldNodeId: op.fieldNodeId, text: op.result }],
    },
    options
  );
  if (!refreshed.ok) return refreshed;
  return {
    ...refreshed,
    effect: { ...refreshed.effect, dirty: [op.paragraphId], dependencyKeys: TEXT_DEPS },
  };
}

/** The `w:instrText` elements between a field's `begin` and its `separate`, in order. */
function instructionNodesOf(
  part: OoxmlPart,
  paragraphId: string,
  fieldNodeId: string
): readonly OoxmlElement[] {
  const paragraph = findNode(part, paragraphId);
  if (!paragraph || paragraph.kind === 'textValue') return [];
  const inOrder: OoxmlElement[] = [];
  const walk = (node: OoxmlNode): void => {
    if (node.kind === 'textValue') return;
    inOrder.push(node);
    node.children.forEach(walk);
  };
  walk(paragraph);
  const begin = inOrder.findIndex((node) => node.id === fieldNodeId);
  if (begin < 0) return [];
  const nodes: OoxmlElement[] = [];
  for (const node of inOrder.slice(begin + 1)) {
    if (isFldCharNode(node)) return nodes;
    if (isInstrText(node)) nodes.push(node);
  }
  return [];
}
