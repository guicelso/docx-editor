// Lowering a BLOCK content control into the line (store lane): the inverse of `liftContentControl`.
//
// A block control holding one paragraph is replaced by that paragraph, and the paragraph holds the
// control, which holds what the paragraph held. The control and the paragraph keep their identities
// and their properties, and no node is copied: only the wrapper changes level. Joining the paragraph
// with its neighbours is not this op's business: `joinParagraphs` does it, by the ids that stay.

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
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { isInlineControl } from './content-control-checkbox.ts';
import { contentControlLockAt } from './tree-op-content-controls.ts';
import { contentControlContentOf, TEXT_DEPS } from './tree-op-nodes.ts';
import type { RevisionAttributionInput, TreeOpRejection, TreeOpResult } from './tree-op-types.ts';

/**
 * Lower a block control holding exactly one paragraph into that paragraph. @public
 *
 * The paragraph takes the control's place among the blocks, keeping its identity and properties,
 * and holds the control, which keeps its identity, its properties and every node the paragraph
 * held. A control holding anything but one paragraph, or one already in a line, is `unsupported`.
 */
export interface LowerContentControlOp {
  readonly op: 'lowerContentControl';
  readonly controlId: string;
  /** A tracked lowering has no implementation yet: an attributed one is refused. */
  readonly revision?: RevisionAttributionInput;
}

export function validateLowerContentControl(
  part: OoxmlPart,
  op: LowerContentControlOp
): TreeOpRejection | null {
  const planned = loweringOf(part, op);
  return typeof planned === 'string' ? planned : null;
}

interface PlannedLowering {
  readonly control: OoxmlElement;
  readonly content: OoxmlElement;
  readonly paragraph: OoxmlElement;
  readonly holder: OoxmlElement;
}

function loweringOf(part: OoxmlPart, op: LowerContentControlOp): PlannedLowering | TreeOpRejection {
  if (op.revision !== undefined) return 'invalidArgs';
  const control = typeof op.controlId === 'string' ? findNode(part, op.controlId) : null;
  if (!control) return 'unknown-content-control';
  if (control.kind !== 'contentControl') return 'not-a-content-control';
  const content = contentControlContentOf(control);
  const holder = parentNodeOf(part, control.id);
  if (isInlineControl(part, control.id) || !content || !holder) return 'unsupported';
  const [paragraph, ...rest] = content.children;
  if (paragraph?.kind !== 'paragraph' || rest.length > 0) return 'unsupported';
  if (contentControlPropertiesOf(control).dataBinding) return 'bound';
  const lock = contentControlLockAt(part, control.id);
  if (lockForbidsEdit(lock) || lockForbidsRemoval(lock)) return 'locked';
  return { control, content, paragraph, holder };
}

export function applyLowerContentControl(
  part: OoxmlPart,
  op: LowerContentControlOp,
  options?: EditOptions
): TreeOpResult {
  const planned = loweringOf(part, op);
  if (typeof planned === 'string') return { ok: false, reason: planned };
  const { control, content, paragraph, holder } = planned;
  const properties = paragraph.children.filter((child) => child.kind === 'paragraphProperties');
  const held = paragraph.children.filter((child) => child.kind !== 'paragraphProperties');
  const inline = {
    ...control,
    children: control.children.map((child) =>
      child.id === content.id ? ({ ...content, children: held } as OoxmlNode) : child
    ),
  } as OoxmlElement;
  const lowered = { ...paragraph, children: [...properties, inline] } as OoxmlElement;
  const children = holder.children.map((child) => (child.id === control.id ? lowered : child));
  const edit = replaceChildren(part, holder.id, children, options);
  if (!edit.ok) return { ok: false, reason: 'tree-invariant', detail: JSON.stringify(edit.issues) };
  return {
    ok: true,
    part: edit.part,
    effect: {
      dirty: [paragraph.id],
      created: [],
      deleted: [],
      dependencyKeys: TEXT_DEPS,
      impact: 'flow-structural',
    },
  };
}
