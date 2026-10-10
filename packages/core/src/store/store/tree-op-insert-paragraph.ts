// A new paragraph at a place between blocks (`tree-op-block-place.ts`): what Enter and typing do
// in the slot outside a block control's tag, where no inline place exists.
//
// It takes the paragraph properties of the edge it stands beside — the last paragraph of the
// block it follows, or the first of the one it precedes — and the face of the run at that edge,
// as a split takes the ones of the paragraph it divides. The section mark stays where it was:
// it ends a section, and a copy would start another one.

import {
  createNodeIdAllocator,
  findNode,
  replaceChildren,
  type EditOptions,
} from '../package/ooxml-edit.ts';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { actorScopedSeed } from '../package/actor-scoped-ids.ts';
import { usedParaIds, w14PrefixInScopeAt, withFreshParaIds } from '../package/para-id.ts';
import { paragraphIdsUnder } from './content-control-value-content.ts';
import {
  blockLandingOf,
  blockPlaceAnchorOf,
  isInsidePlace,
  type BlockLanding,
  type BlockPlace,
} from './tree-op-block-place.ts';
import { cloneWithFreshIds, textRun, wmlElement } from './tree-op-content-controls.ts';
import { isInsertableText } from './tree-op-inline-elements.ts';
import { paragraphPropertiesNodeOf, runPropertiesNodeOf, TEXT_DEPS } from './tree-op-nodes.ts';
import { withoutSectionMark } from './tree-op-section.ts';
import type { RevisionAttributionInput, TreeOpRejection, TreeOpResult } from './tree-op-types.ts';

/** A new paragraph before or after a block, holding `text` when given. The caret goes to it. @public */
export interface InsertParagraphOp {
  readonly op: 'insertParagraph';
  /** Before or after a block. `inside` a control is refused: its prompt takes typing as is. */
  readonly at: BlockPlace;
  /** Text the paragraph holds: one line, never empty. Absent writes an empty paragraph. */
  readonly text?: string;
  /** A tracked paragraph has no implementation yet: an attributed insertion is refused. */
  readonly revision?: RevisionAttributionInput;
}

export function validateInsertParagraph(
  part: OoxmlPart,
  op: InsertParagraphOp
): TreeOpRejection | null {
  if (op.revision !== undefined) return 'invalidArgs';
  if (op.text !== undefined && !isWritableLine(op.text)) return 'invalid-text';
  if (blockPlaceAnchorOf(op.at) === null || isInsidePlace(op.at)) return 'invalidArgs';
  const landing = blockLandingOf(part, op.at);
  return typeof landing === 'string' ? landing : null;
}

export function applyInsertParagraph(
  part: OoxmlPart,
  op: InsertParagraphOp,
  options?: EditOptions
): TreeOpResult {
  const refused = validateInsertParagraph(part, op);
  if (refused) return { ok: false, reason: refused };
  const landing = blockLandingOf(part, op.at) as BlockLanding;
  const side = 'before' in op.at ? 'first' : 'last';
  const edge = edgeParagraphOf(part, blockPlaceAnchorOf(op.at)!, side);
  const nextId = createNodeIdAllocator(part);
  const paragraph = withNewIdentity(part, landing, newParagraph(edge, side, op.text, nextId));
  const children = landing.holder.children;
  const edit = replaceChildren(
    part,
    landing.holder.id,
    [...children.slice(0, landing.index), paragraph, ...children.slice(landing.index)],
    options
  );
  if (!edit.ok) return { ok: false, reason: 'tree-invariant', detail: JSON.stringify(edit.issues) };
  return {
    ok: true,
    part: edit.part,
    effect: {
      dirty: [landing.holder.id],
      created: [paragraph.id],
      deleted: [],
      caret: { paragraphId: paragraph.id },
      dependencyKeys: TEXT_DEPS,
      impact: 'flow-structural',
    },
  };
}

function isWritableLine(text: unknown): boolean {
  return typeof text === 'string' && text.length > 0 && isInsertableText(text);
}

/** The paragraph at the edge of a block the new one stands beside, or null for none. */
function edgeParagraphOf(
  part: OoxmlPart,
  anchorId: string,
  side: 'first' | 'last'
): OoxmlElement | null {
  const anchor = findNode(part, anchorId);
  if (!anchor) return null;
  const paragraphs = paragraphIdsUnder(anchor);
  const id = side === 'first' ? paragraphs[0] : paragraphs[paragraphs.length - 1];
  const paragraph = id === undefined ? null : findNode(part, id);
  return paragraph?.kind === 'paragraph' ? paragraph : null;
}

function newParagraph(
  edge: OoxmlElement | null,
  side: 'first' | 'last',
  text: string | undefined,
  nextId: () => string
): OoxmlNode {
  const properties = edge ? paragraphPropertiesNodeOf(edge) : undefined;
  const kept = properties ? withoutSectionMark(properties) : undefined;
  const face = edge && text !== undefined ? edgeRunPropertiesOf(edge, side) : undefined;
  const children = [
    ...(kept ? [cloneWithFreshIds(kept, nextId)] : []),
    ...(text === undefined ? [] : [textRun(nextId, text, face && cloneWithFreshIds(face, nextId))]),
  ];
  return wmlElement(nextId, 'p', { kind: 'paragraph' as OoxmlNode['kind'], children });
}

/** The face of the run at the paragraph's edge: its last run after it, its first before it. */
function edgeRunPropertiesOf(
  paragraph: OoxmlElement,
  side: 'first' | 'last'
): OoxmlNode | undefined {
  const runs = paragraph.children.filter((child) => child.kind === 'run');
  const run = side === 'first' ? runs[0] : runs[runs.length - 1];
  return run && run.kind !== 'textValue' ? (runPropertiesNodeOf(run) ?? undefined) : undefined;
}

function withNewIdentity(part: OoxmlPart, landing: BlockLanding, paragraph: OoxmlNode): OoxmlNode {
  return withFreshParaIds(
    paragraph,
    new Set(usedParaIds(part.root as OoxmlElement)),
    actorScopedSeed(`${landing.holder.id}:${landing.index}:paragraph`),
    { value: 0 },
    w14PrefixInScopeAt(part, landing.holder)
  );
}
