// Lifting an INLINE content control to the block level (store lane).
//
// The paragraph holding the control is cut at the control's edges, through the same split a
// paragraph mark makes: the text before it stays in the paragraph, the text after it goes to a new
// one. The paragraph left holding only the control becomes the control's one paragraph, with the
// control's content as its own. The control keeps its identity and its properties, and no node it
// holds is copied: only the wrapper around them changes level. A side with nothing on it is not cut,
// so no empty paragraph is left behind.

import {
  contentControlPropertiesOf,
  lockForbidsEdit,
  lockForbidsRemoval,
} from '../package/content-control-nodes.ts';
import {
  findNode,
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
import { isInlineControl } from './content-control-checkbox.ts';
import { applySplit } from './tree-op-apply.ts';
import { contentControlLockAt } from './tree-op-content-controls.ts';
import { contentControlContentOf, TEXT_DEPS } from './tree-op-nodes.ts';
import { paragraphOffsetIndex } from './tree-op-segments.ts';
import type { RevisionAttributionInput, TreeOpRejection, TreeOpResult } from './tree-op-types.ts';

/**
 * Lift an inline control that a paragraph holds directly to a block control in its place. @public
 *
 * The paragraph is cut at the control's edges: what comes before stays in it, what comes after
 * goes to a new paragraph with the same properties, and a side with nothing on it is not cut. The
 * control keeps its identity, its properties and every node it holds, now in one paragraph with
 * the properties of the one it came from. A control held by another inline control or a
 * hyperlink, or one already between blocks, is `unsupported`.
 */
export interface LiftContentControlOp {
  readonly op: 'liftContentControl';
  readonly controlId: string;
  /** A tracked lift has no implementation yet: an attributed one is refused. */
  readonly revision?: RevisionAttributionInput;
}

export function validateLiftContentControl(
  part: OoxmlPart,
  op: LiftContentControlOp
): TreeOpRejection | null {
  const planned = liftOf(part, op);
  return typeof planned === 'string' ? planned : null;
}

interface PlannedLift {
  readonly controlId: string;
  readonly paragraphId: string;
  readonly start: number;
  readonly end: number;
  readonly length: number;
}

function liftOf(part: OoxmlPart, op: LiftContentControlOp): PlannedLift | TreeOpRejection {
  if (op.revision !== undefined) return 'invalidArgs';
  const control = typeof op.controlId === 'string' ? findNode(part, op.controlId) : null;
  if (!control) return 'unknown-content-control';
  if (control.kind !== 'contentControl') return 'not-a-content-control';
  if (!isInlineControl(part, control.id) || !contentControlContentOf(control)) return 'unsupported';
  const paragraph = parentNodeOf(part, control.id);
  if (paragraph?.kind !== 'paragraph') return 'unsupported';
  if (contentControlPropertiesOf(control).dataBinding) return 'bound';
  const lock = contentControlLockAt(part, control.id);
  if (lockForbidsEdit(lock) || lockForbidsRemoval(lock)) return 'locked';
  const index = paragraphOffsetIndex(paragraph);
  const span = index.spanOf(control);
  if (!span) return 'tree-invariant';
  return {
    controlId: control.id,
    paragraphId: paragraph.id,
    start: span.start,
    end: span.end,
    length: index.length,
  };
}

export function applyLiftContentControl(
  part: OoxmlPart,
  op: LiftContentControlOp,
  options?: EditOptions
): TreeOpResult {
  const planned = liftOf(part, op);
  if (typeof planned === 'string') return { ok: false, reason: planned };
  const afterCut = cutAt(
    part,
    planned.paragraphId,
    planned.end < planned.length ? planned.end : null,
    options
  );
  if (!afterCut.ok) return afterCut;
  const beforeCut = cutAt(
    afterCut.part,
    planned.paragraphId,
    planned.start > 0 ? planned.start : null,
    options
  );
  if (!beforeCut.ok) return beforeCut;
  return lifted(
    beforeCut.part,
    beforeCut.holding,
    planned,
    [...afterCut.created, ...beforeCut.created],
    options
  );
}

/** `holding` is the paragraph from the cut on: the new tail, or the paragraph itself when there is no cut. */
type Cut =
  | Readonly<{ ok: true; part: OoxmlPart; holding: string; created: readonly string[] }>
  | Readonly<{ ok: false; reason: TreeOpRejection; detail?: string }>;

/** The paragraph cut at an offset: it keeps its identity, and the tail is the new paragraph. */
function cutAt(
  part: OoxmlPart,
  paragraphId: string,
  offset: number | null,
  options?: EditOptions
): Cut {
  if (offset === null) return { ok: true, part, holding: paragraphId, created: [] };
  const paragraph = findNode(part, paragraphId) as OoxmlParagraphNode;
  const split = applySplit(part, paragraph, offset, options);
  if (!split.ok) return split;
  const tail = split.effect.split?.tail;
  if (tail === undefined)
    return { ok: false, reason: 'tree-invariant', detail: 'a split named no tail' };
  return { ok: true, part: split.part, holding: tail, created: [tail] };
}

/** The paragraph that now holds the control becomes its one paragraph, in the control's place. */
function lifted(
  part: OoxmlPart,
  paragraphId: string,
  planned: PlannedLift,
  created: readonly string[],
  options?: EditOptions
): TreeOpResult {
  const paragraph = findNode(part, paragraphId);
  const control = findNode(part, planned.controlId);
  const holder = parentNodeOf(part, paragraphId);
  const content = control ? contentControlContentOf(control) : undefined;
  if (
    !paragraph ||
    paragraph.kind !== 'paragraph' ||
    control?.kind !== 'contentControl' ||
    !content ||
    !holder
  ) {
    return { ok: false, reason: 'tree-invariant', detail: 'the lifted control left its paragraph' };
  }
  const own = {
    ...paragraph,
    children: paragraph.children.flatMap((child) =>
      child.id === control.id ? content.children : [child]
    ),
  };
  const block = {
    ...control,
    children: control.children.map((child) =>
      child.id === content.id ? ({ ...content, children: [own] } as OoxmlNode) : child
    ),
  } as OoxmlElement;
  const children = holder.children.map((child) => (child.id === paragraphId ? block : child));
  const edit = replaceChildren(part, holder.id, children, options);
  if (!edit.ok) return { ok: false, reason: 'tree-invariant', detail: JSON.stringify(edit.issues) };
  return {
    ok: true,
    part: edit.part,
    effect: {
      dirty: [planned.paragraphId, ...created],
      created: [...created],
      deleted: [],
      dependencyKeys: TEXT_DEPS,
      impact: 'flow-structural',
    },
  };
}
