// Writing at the BLOCK level (store lane): sibling blocks wrapped whole in a new control, or a new
// control at a place between blocks (`tree-op-block-place.ts`).
//
// A wrap is addressed by the blocks it takes, a new control by the place it goes, and each
// refuses for its own reasons. The controls are `w:sdt` around blocks, the `CT_SdtBlock`
// of §17.5.2.29; unwrapping either is `removeContentControl`, which already splices a control's
// children into its place at any level.

import type { ContentControlLock } from '../package/content-control-nodes.ts';
import { actorScopedSeed } from '../package/actor-scoped-ids.ts';
import {
  createNodeIdAllocator,
  parentNodeOf,
  replaceChildren,
  type EditOptions,
} from '../package/ooxml-edit.ts';
import type {
  OoxmlElement,
  OoxmlNode,
  OoxmlParagraphNode,
  OoxmlPart,
} from '../package/ooxml-tree.ts';
import { usedParaIds, w14PrefixInScopeAt, withFreshParaIds } from '../package/para-id.ts';
import {
  blockLandingOf,
  blockOf,
  cellKeepsItsLastBlock,
  holdsPrompt,
  type BlockLanding,
  type BlockPlace,
} from './tree-op-block-place.ts';
import { controlElement, propertiesFor } from './tree-op-content-control-insert.ts';
import {
  contentControlEffect,
  contentWithText,
  editedProperties,
  isWritableContentControlMetadata,
  promptFor,
} from './tree-op-content-controls.ts';
import {
  fragmentShape,
  MAX_FRAGMENT_INSERT_BLOCKS,
  withRequiredNamespaceBindings,
} from './tree-op-fragment.ts';
import { cloneWithNewIds, contentControlContentOf } from './tree-op-nodes.ts';
import type {
  RevisionAttributionInput,
  TreeDocOp,
  TreeOpRejection,
  TreeOpResult,
} from './tree-op-types.ts';
import { CONTENT_CONTROL_LOCKS } from './tree-op-validate-controls.ts';
import { contentControlPropertiesContainerOf } from '../package/content-control-nodes.ts';

/**
 * Wrap sibling blocks — paragraphs and block-level controls — in a new rich-text control.
 *
 * Addressed by the BLOCKS, not by paragraphs inside them: naming a control as both edges wraps
 * the control itself, naming its first and last child wraps its content. A table is refused,
 * since a tag drawn at its edge would have no line to stand on. @public
 */
export interface WrapBlocksInContentControlOp {
  readonly op: 'wrapBlocksInContentControl';
  readonly firstBlockId: string;
  readonly lastBlockId: string;
  readonly tag?: string;
  readonly lock?: ContentControlLock;
  /** A tracked block control has no implementation yet: an attributed wrap is refused. */
  readonly revision?: RevisionAttributionInput;
}

/**
 * A new rich-text block control at a place between blocks. @public
 *
 * With `paragraphs` it holds them, each with fresh node and paragraph ids; without, it holds one
 * paragraph showing the prompt, which the first keystroke replaces.
 */
export interface InsertBlockContentControlOp {
  readonly op: 'insertBlockContentControl';
  readonly at: BlockPlace;
  readonly tag?: string;
  readonly lock?: ContentControlLock;
  readonly paragraphs?: readonly OoxmlParagraphNode[];
  /** A tracked block control has no implementation yet: an attributed insertion is refused. */
  readonly revision?: RevisionAttributionInput;
}

type BlockControlOp = WrapBlocksInContentControlOp | InsertBlockContentControlOp;

/** The ops that write at the block level. */
export type BlockStructureOp = BlockControlOp;

export function isBlockStructureOp(op: TreeDocOp): op is BlockStructureOp {
  return op.op === 'wrapBlocksInContentControl' || op.op === 'insertBlockContentControl';
}

export function validateBlockStructureOp(
  part: OoxmlPart,
  op: BlockStructureOp
): TreeOpRejection | null {
  return validateBlockControlOp(part, op);
}

export function applyBlockStructureOp(
  part: OoxmlPart,
  op: BlockStructureOp,
  options?: EditOptions
): TreeOpResult {
  return applyBlockControlOp(part, op, options);
}

function validateBlockControlOp(part: OoxmlPart, op: BlockControlOp): TreeOpRejection | null {
  if (op.revision !== undefined) return 'invalidArgs';
  if (!isWritableContentControlMetadata(op.tag)) return 'invalid-property-value';
  if (op.lock !== undefined && !CONTENT_CONTROL_LOCKS.has(op.lock)) return 'invalidArgs';
  if (op.op === 'wrapBlocksInContentControl') {
    const wrap = wrappedBlocks(part, op);
    return typeof wrap === 'string' ? wrap : null;
  }
  const landing = blockLandingOf(part, op.at);
  if (typeof landing === 'string') return landing;
  if (op.paragraphs !== undefined) {
    const refused = paragraphsRefusal(op.paragraphs);
    if (refused) return refused;
  }
  return cellKeepsItsLastBlock(landing.holder, landing.index) ? null : 'block-required';
}

function applyBlockControlOp(
  part: OoxmlPart,
  op: BlockControlOp,
  options?: EditOptions
): TreeOpResult {
  const refused = validateBlockControlOp(part, op);
  if (refused) return { ok: false, reason: refused };
  return op.op === 'wrapBlocksInContentControl'
    ? applyWrap(part, op, options)
    : applyInsert(part, op, options);
}

interface WrappedBlocks {
  readonly holder: OoxmlElement;
  readonly first: number;
  readonly last: number;
}

/** The siblings a wrap takes, in order, or why it takes none. */
function wrappedBlocks(
  part: OoxmlPart,
  op: WrapBlocksInContentControlOp
): WrappedBlocks | TreeOpRejection {
  const first = blockOf(part, op.firstBlockId);
  if (typeof first === 'string') return first;
  const last = blockOf(part, op.lastBlockId);
  if (typeof last === 'string') return last;
  const holder = parentNodeOf(part, first.id);
  if (!holder) return 'not-a-block';
  if (parentNodeOf(part, last.id)?.id !== holder.id) return 'not-adjacent-siblings';
  const from = holder.children.findIndex((child) => child.id === first.id);
  const to = holder.children.findIndex((child) => child.id === last.id);
  if (from > to) return 'not-adjacent-siblings';
  if (holder.children.slice(from, to + 1).some((child) => child.kind === 'table')) {
    return 'unsupported';
  }
  if (holdsPrompt(part, holder)) return 'invalidArgs';
  return cellKeepsItsLastBlock(holder, to + 1)
    ? { holder, first: from, last: to }
    : 'block-required';
}

function applyWrap(
  part: OoxmlPart,
  op: WrapBlocksInContentControlOp,
  options?: EditOptions
): TreeOpResult {
  const wrapped = wrappedBlocks(part, op);
  if (typeof wrapped === 'string') return { ok: false, reason: wrapped };
  const nextId = createNodeIdAllocator(part);
  const content = wrapped.holder.children.slice(wrapped.first, wrapped.last + 1);
  const control = controlElement(newControlProperties(part, op, false, nextId), content, nextId);
  return written(part, wrapped.holder, wrappedChildren(wrapped, control), options);
}

function applyInsert(
  part: OoxmlPart,
  op: InsertBlockContentControlOp,
  options?: EditOptions
): TreeOpResult {
  const landing = blockLandingOf(part, op.at);
  if (typeof landing === 'string') return { ok: false, reason: landing };
  const nextId = createNodeIdAllocator(part);
  const prompted = op.paragraphs === undefined;
  const authored = prompted
    ? contentWithText(undefined, promptFor('richText', options), nextId, false)
    : op.paragraphs.map((paragraph) => cloneWithNewIds(paragraph, nextId));
  if (!authored) return { ok: false, reason: 'unsupported' };
  const content = withNewIdentities(part, landing, authored);
  const control = controlElement(newControlProperties(part, op, prompted, nextId), content, nextId);
  const bound = withRequiredNamespaceBindings(part, content);
  if (landing.prompt === null) {
    return written(bound, landing.holder, landed(landing, [control]), options);
  }
  return written(bound, landing.prompt, promptReplacedBy(landing.prompt, control, nextId), options);
}

function newControlProperties(
  part: OoxmlPart,
  op: BlockControlOp,
  showingPlaceholder: boolean,
  nextId: () => string
): OoxmlElement {
  return propertiesFor(
    part,
    {
      type: 'richText',
      showingPlaceholder,
      ...(op.tag === undefined ? {} : { tag: op.tag }),
      ...(op.lock === undefined ? {} : { lock: op.lock }),
    },
    nextId
  );
}

/** A control showing its prompt, whose content is now `control` and whose flag is cleared. */
function promptReplacedBy(
  owner: OoxmlElement,
  control: OoxmlNode,
  nextId: () => string
): readonly OoxmlNode[] {
  const sdtPr = contentControlPropertiesContainerOf(owner);
  const content = contentControlContentOf(owner)!;
  const properties = editedProperties(sdtPr, { showingPlaceholder: false }, nextId);
  return owner.children.map((child) => {
    if (child.id === sdtPr?.id) return properties;
    return child.id === content.id ? ({ ...content, children: [control] } as OoxmlNode) : child;
  });
}

/**
 * Every paragraph the op writes gets a fresh `w14:paraId`: the given ones carry ids that belong
 * elsewhere, and the prompt's has none, which would leave it unaddressable.
 */
function withNewIdentities(
  part: OoxmlPart,
  landing: BlockLanding,
  paragraphs: readonly OoxmlNode[]
): readonly OoxmlNode[] {
  const paraIds = new Set(usedParaIds(part.root as OoxmlElement));
  const counter = { value: 0 };
  const prefix = w14PrefixInScopeAt(part, landing.holder);
  return paragraphs.map((paragraph, index) =>
    withFreshParaIds(
      paragraph,
      paraIds,
      actorScopedSeed(`${landing.holder.id}:${landing.index}:block-control:${index}`),
      counter,
      prefix
    )
  );
}

function paragraphsRefusal(paragraphs: readonly OoxmlParagraphNode[]): TreeOpRejection | null {
  if (!Array.isArray(paragraphs) || paragraphs.length === 0) return 'invalidArgs';
  if (paragraphs.length > MAX_FRAGMENT_INSERT_BLOCKS) return 'fragment-resource-budget';
  const budget = { nodes: 0 };
  for (const paragraph of paragraphs) {
    if (paragraph?.kind !== 'paragraph') return 'fragment-invalid-block';
    const refused = fragmentShape(paragraph, 1, budget);
    if (refused) return refused;
  }
  return null;
}

function landed(landing: BlockLanding, blocks: readonly OoxmlNode[]): readonly OoxmlNode[] {
  if (landing.prompt !== null) return blocks;
  const children = landing.holder.children;
  return [...children.slice(0, landing.index), ...blocks, ...children.slice(landing.index)];
}

function wrappedChildren(wrapped: WrappedBlocks, control: OoxmlNode): readonly OoxmlNode[] {
  const children = wrapped.holder.children;
  return [...children.slice(0, wrapped.first), control, ...children.slice(wrapped.last + 1)];
}

function written(
  part: OoxmlPart,
  holder: OoxmlElement,
  children: readonly OoxmlNode[],
  options?: EditOptions
): TreeOpResult {
  const edit = replaceChildren(part, holder.id, children, options);
  if (!edit.ok) return { ok: false, reason: 'tree-invariant', detail: JSON.stringify(edit.issues) };
  return { ok: true, part: edit.part, effect: contentControlEffect(holder.id, 'flow-structural') };
}
