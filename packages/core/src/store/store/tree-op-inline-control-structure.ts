// Dividing and joining INLINE content controls (store lane): the inline twins of `splitParagraph`
// and `joinParagraphs`. Lifting one to the block level, and lowering a block control into the line,
// are the other two ops of the family, in `tree-op-control-lift.ts` and `tree-op-control-lower.ts`.
//
// A split cuts one control in two siblings at a place strictly inside its content, the way a
// paragraph mark cuts a paragraph — for a block control, a place in one of the paragraphs it holds:
// the head keeps the control's identity, and the tail is a new control with the tag the caller
// names and a fresh `w:id`. Copying the tag, which is what a
// paragraph split does to a control it crosses, would leave two controls claiming one identity.
// A join puts the second control's content at the end of the first and removes the second.
// Neither changes a character: only the wrappers around them move.

import {
  allocateContentControlId,
  contentControlPropertiesContainerOf,
  contentControlPropertiesOf,
  lockForbidsEdit,
  lockForbidsRemoval,
} from '../package/content-control-nodes.ts';
import {
  createNodeIdAllocator,
  findNode,
  parentNodeOf,
  replaceChildren,
  type EditOptions,
} from '../package/ooxml-edit.ts';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { isInlineControl } from './content-control-checkbox.ts';
import { applyTreeOp, distributeInline } from './tree-op-apply.ts';
import {
  applyLiftContentControl,
  validateLiftContentControl,
  type LiftContentControlOp,
} from './tree-op-control-lift.ts';
import {
  applyLowerContentControl,
  validateLowerContentControl,
  type LowerContentControlOp,
} from './tree-op-control-lower.ts';
import {
  contentControlEffect,
  contentControlLockAt,
  editedProperties,
  isWritableContentControlMetadata,
} from './tree-op-content-controls.ts';
import {
  cloneWithNewIds,
  contentControlContentOf,
  isShowingPlaceholder,
  paragraphContainingNode,
  TEXT_DEPS,
} from './tree-op-nodes.ts';
import { paragraphOffsetIndex, segmentsOf, splitsSurrogate } from './tree-op-segments.ts';
import type {
  RevisionAttributionInput,
  TreeDocOp,
  TreeOpEffect,
  TreeOpRejection,
  TreeOpResult,
} from './tree-op-types.ts';

/**
 * Cut a content control in two at a place strictly inside its content: an inline one at a
 * paragraph `offset`, a block one `at` an offset in a paragraph it holds. @public
 *
 * The head keeps the control and everything it declares; the tail is a new control with the
 * same properties, `tag` and a fresh `w:id`. The place must be in the control's own content:
 * inside a control nested in it, a hyperlink, a revision wrapper or an atomic field is
 * `indivisible-content`, since dividing those would copy their identity too. In a block control,
 * the start of a paragraph moves that paragraph to the tail and its end moves the blocks after it,
 * whole and with their ids; a place inside a paragraph cuts it, as `splitParagraph` does.
 */
export type SplitContentControlOp = {
  readonly op: 'splitContentControl';
  readonly controlId: string;
  readonly tag: string;
  /** A tracked split has no implementation yet: an attributed one is refused. */
  readonly revision?: RevisionAttributionInput;
} & (
  | { readonly offset: number; readonly at?: undefined }
  | { readonly at: BlockControlPlace; readonly offset?: undefined }
);

/** A place in a paragraph a block control holds as one of its own blocks. @public */
export interface BlockControlPlace {
  readonly paragraphId: string;
  readonly offset: number;
}

/**
 * Join two adjacent sibling inline controls: the second's content goes to the end of the first,
 * and the second goes. A control showing its prompt holds nothing to keep. @public
 */
export interface JoinContentControlsOp {
  readonly op: 'joinContentControls';
  readonly firstId: string;
  readonly secondId: string;
  /** A tracked join has no implementation yet: an attributed one is refused. */
  readonly revision?: RevisionAttributionInput;
}

/** The ops that divide, join, lift or lower controls between the line and the blocks. */
export type InlineControlOp =
  | SplitContentControlOp
  | JoinContentControlsOp
  | LiftContentControlOp
  | LowerContentControlOp;

export function isInlineControlOp(op: TreeDocOp): op is InlineControlOp {
  return (
    op.op === 'splitContentControl' ||
    op.op === 'joinContentControls' ||
    op.op === 'liftContentControl' ||
    op.op === 'lowerContentControl'
  );
}

export function validateInlineControlOp(
  part: OoxmlPart,
  op: InlineControlOp
): TreeOpRejection | null {
  if (op.op === 'liftContentControl') return validateLiftContentControl(part, op);
  if (op.op === 'lowerContentControl') return validateLowerContentControl(part, op);
  const planned = op.op === 'splitContentControl' ? splitPlanOf(part, op) : joinOf(part, op);
  return typeof planned === 'string' ? planned : null;
}

export function applyInlineControlOp(
  part: OoxmlPart,
  op: InlineControlOp,
  options?: EditOptions
): TreeOpResult {
  switch (op.op) {
    case 'splitContentControl':
      return op.offset === undefined
        ? applyBlockSplit(part, op, op.at, options)
        : applySplit(part, op, options);
    case 'joinContentControls':
      return applyJoin(part, op, options);
    case 'liftContentControl':
      return applyLiftContentControl(part, op, options);
    case 'lowerContentControl':
      return applyLowerContentControl(part, op, options);
  }
}

interface PlacedControl {
  readonly control: OoxmlElement;
  readonly holder: OoxmlElement;
  readonly paragraphId: string;
}

interface PlannedSplit extends PlacedControl {
  readonly offset: number;
}

interface PlannedJoin {
  readonly first: PlacedControl;
  readonly second: OoxmlElement;
}

interface PlannedBlockSplit {
  readonly control: OoxmlElement;
  readonly holder: OoxmlElement;
  /** The first block the tail takes, among the control's content, once a cut paragraph is cut. */
  readonly index: number;
  /** The paragraph cut in two at the place, or null at its start or its end. */
  readonly cut: BlockControlPlace | null;
}

type InlineSplitOp = Extract<SplitContentControlOp, { readonly offset: number }>;

function splitPlanOf(
  part: OoxmlPart,
  op: SplitContentControlOp
): PlannedSplit | PlannedBlockSplit | TreeOpRejection {
  return op.offset === undefined ? blockSplitOf(part, op.controlId, op.at, op) : splitOf(part, op);
}

function splitOf(part: OoxmlPart, op: InlineSplitOp): PlannedSplit | TreeOpRejection {
  if (op.revision !== undefined) return 'invalidArgs';
  if (!isWritableContentControlMetadata(op.tag) || typeof op.tag !== 'string') {
    return 'invalid-property-value';
  }
  const placed = placedInlineControl(part, op.controlId);
  if (typeof placed === 'string') return placed;
  if (isShowingPlaceholder(placed.control)) return 'invalidArgs';
  const refused = editRefusal(part, placed.control, true);
  if (refused) return refused;
  const paragraph = paragraphContainingNode(part, placed.control.id)!;
  const span = paragraphOffsetIndex(paragraph).spanOf(placed.control);
  if (!Number.isInteger(op.offset) || !span || op.offset <= span.start || op.offset >= span.end) {
    return 'invalid-range';
  }
  if (splitsSurrogate(paragraph, op.offset)) return 'splits-surrogate-pair';
  if (indivisibleInside(contentControlContentOf(placed.control)!.children, paragraph, op.offset)) {
    return 'indivisible-content';
  }
  return { ...placed, offset: op.offset };
}

function blockSplitOf(
  part: OoxmlPart,
  controlId: string,
  at: BlockControlPlace,
  op: SplitContentControlOp
): PlannedBlockSplit | TreeOpRejection {
  if (op.revision !== undefined) return 'invalidArgs';
  if (!isWritableContentControlMetadata(op.tag) || typeof op.tag !== 'string') {
    return 'invalid-property-value';
  }
  const control = typeof controlId === 'string' ? findNode(part, controlId) : null;
  if (!control) return 'unknown-content-control';
  if (control.kind !== 'contentControl') return 'not-a-content-control';
  const content = contentControlContentOf(control);
  if (isInlineControl(part, control.id) || !content) return 'unsupported';
  if (isShowingPlaceholder(control)) return 'invalidArgs';
  const holder = parentNodeOf(part, control.id);
  if (!holder) return 'tree-invariant';
  const blocks = content.children;
  const index = blocks.findIndex((child) => child.id === at?.paragraphId);
  const paragraph = blocks[index];
  if (!paragraph || paragraph.kind !== 'paragraph') return 'invalid-range';
  const length = paragraphOffsetIndex(paragraph).length;
  if (!Number.isInteger(at.offset) || at.offset < 0 || at.offset > length) return 'invalid-range';
  if (index === 0 && at.offset === 0) return 'invalid-range';
  if (index === blocks.length - 1 && at.offset === length) return 'invalid-range';
  const refused = editRefusal(part, control, true);
  if (refused) return refused;
  if (at.offset === 0) return { control, holder, index, cut: null };
  if (at.offset === length) return { control, holder, index: index + 1, cut: null };
  if (splitsSurrogate(paragraph, at.offset)) return 'splits-surrogate-pair';
  if (indivisibleInside(paragraph.children, paragraph, at.offset)) return 'indivisible-content';
  return { control, holder, index: index + 1, cut: at };
}

function joinOf(part: OoxmlPart, op: JoinContentControlsOp): PlannedJoin | TreeOpRejection {
  if (op.revision !== undefined) return 'invalidArgs';
  const first = placedInlineControl(part, op.firstId);
  if (typeof first === 'string') return first;
  const second = placedInlineControl(part, op.secondId);
  if (typeof second === 'string') return second;
  const siblings = first.holder.children;
  const at = siblings.findIndex((child) => child.id === first.control.id);
  if (second.holder.id !== first.holder.id || siblings[at + 1]?.id !== second.control.id) {
    return 'not-adjacent-siblings';
  }
  return (
    editRefusal(part, first.control, false) ??
    editRefusal(part, second.control, true) ?? {
      first,
      second: second.control,
    }
  );
}

function placedInlineControl(part: OoxmlPart, controlId: string): PlacedControl | TreeOpRejection {
  const control = typeof controlId === 'string' ? findNode(part, controlId) : null;
  if (!control) return 'unknown-content-control';
  if (control.kind !== 'contentControl') return 'not-a-content-control';
  if (!isInlineControl(part, control.id) || !contentControlContentOf(control)) return 'unsupported';
  const holder = parentNodeOf(part, control.id);
  const paragraph = paragraphContainingNode(part, control.id);
  if (!holder || !paragraph) return 'tree-invariant';
  return { control, holder, paragraphId: paragraph.id };
}

/**
 * The control's own `w:lock` and every enclosing one: the content is rewritten, and a control whose
 * wrapper is removed or minted again answers for its wrapper too.
 */
function editRefusal(
  part: OoxmlPart,
  control: OoxmlElement,
  wrapperChanges: boolean
): TreeOpRejection | null {
  if (contentControlPropertiesOf(control).dataBinding) return 'bound';
  const lock = contentControlLockAt(part, control.id);
  if (lockForbidsEdit(lock)) return 'locked';
  return wrapperChanges && lockForbidsRemoval(lock) ? 'locked' : null;
}

/**
 * Whether the offset falls strictly inside an atomic field, or inside a child of the control's
 * content that is not a run — a nested control, a hyperlink, a revision wrapper.
 */
function indivisibleInside(
  children: readonly OoxmlNode[],
  paragraph: Parameters<typeof paragraphOffsetIndex>[0],
  offset: number
): boolean {
  const index = paragraphOffsetIndex(paragraph);
  const strictlyInside = (start: number, end: number): boolean => start < offset && offset < end;
  if (
    index.segments.some(
      (segment) => segment.removeNodeIds && strictlyInside(segment.start, segment.end)
    )
  ) {
    return true;
  }
  return children.some((child) => {
    if (child.kind === 'run' || child.kind === 'paragraphProperties') return false;
    const span = index.spanOf(child);
    return span !== null && strictlyInside(span.start, span.end);
  });
}

function applySplit(part: OoxmlPart, op: InlineSplitOp, options?: EditOptions): TreeOpResult {
  const planned = splitOf(part, op);
  if (typeof planned === 'string') return { ok: false, reason: planned };
  const allocated = allocateContentControlId(part.root);
  if (allocated === null) return { ok: false, reason: 'resource-limit' };
  const nextId = createNodeIdAllocator(part);
  const paragraph = findNode(part, planned.paragraphId) as Parameters<typeof segmentsOf>[0];
  const [head, tail] = distributeInline(
    planned.control,
    [planned.offset],
    2,
    segmentsOf(paragraph),
    nextId
  );
  const headControl = head?.[0];
  const tailControl = tail?.[0];
  if (!headControl || !tailControl || tailControl.kind === 'textValue') {
    return { ok: false, reason: 'tree-invariant', detail: 'split produced an empty half' };
  }
  const renamed = withProperties(
    tailControl,
    editedProperties(
      contentControlPropertiesContainerOf(tailControl),
      { tag: op.tag, id: allocated },
      nextId
    )
  );
  const children = planned.holder.children.flatMap((child) =>
    child.id === planned.control.id ? [headControl, renamed] : [child]
  );
  return written(
    part,
    planned.holder,
    children,
    effectOf(planned.paragraphId, [renamed.id], []),
    options
  );
}

function applyBlockSplit(
  part: OoxmlPart,
  op: SplitContentControlOp,
  at: BlockControlPlace,
  options?: EditOptions
): TreeOpResult {
  const planned = blockSplitOf(part, op.controlId, at, op);
  if (typeof planned === 'string') return { ok: false, reason: planned };
  const cut =
    planned.cut === null
      ? null
      : applyTreeOp(part, { op: 'splitParagraph', ...planned.cut }, options);
  if (cut !== null && !cut.ok) return cut;
  const current = cut === null ? part : cut.part;
  const allocated = allocateContentControlId(current.root);
  if (allocated === null) return { ok: false, reason: 'resource-limit' };
  const nextId = createNodeIdAllocator(current);
  const control = findNode(current, planned.control.id) as OoxmlElement;
  const holder = findNode(current, planned.holder.id) as OoxmlElement;
  const { index } = planned;
  const content = contentControlContentOf(control)!;
  const head = withContent(control, content.children.slice(0, index), nextId);
  const tail = blockTailOf(
    control,
    content.children.slice(index),
    { tag: op.tag, id: allocated },
    nextId
  );
  const children = holder.children.flatMap((child) =>
    child.id === control.id ? [head, tail] : [child]
  );
  const edit = replaceChildren(current, holder.id, children, options);
  if (!edit.ok) return { ok: false, reason: 'tree-invariant', detail: JSON.stringify(edit.issues) };
  return {
    ok: true,
    part: edit.part,
    effect: {
      ...contentControlEffect(holder.id, 'flow-structural'),
      created: [...(cut === null ? [] : cut.effect.created), tail.id],
    },
  };
}

/** A new control around the moved blocks: the wrapper and its properties minted again, the blocks kept. */
function blockTailOf(
  control: OoxmlElement,
  blocks: readonly OoxmlNode[],
  edits: { readonly tag: string; readonly id: number },
  nextId: () => string
): OoxmlElement {
  const sdtPr = contentControlPropertiesContainerOf(control);
  const content = contentControlContentOf(control)!;
  return {
    ...control,
    id: nextId(),
    children: control.children.map((child) => {
      if (child.id === content.id)
        return { ...content, id: nextId(), children: blocks } as OoxmlNode;
      if (child.id === sdtPr?.id) {
        return editedProperties(cloneWithNewIds(sdtPr, nextId) as OoxmlElement, edits, nextId);
      }
      return cloneWithNewIds(child, nextId);
    }),
  } as OoxmlElement;
}

function applyJoin(
  part: OoxmlPart,
  op: JoinContentControlsOp,
  options?: EditOptions
): TreeOpResult {
  const planned = joinOf(part, op);
  if (typeof planned === 'string') return { ok: false, reason: planned };
  const { first, second } = planned;
  const moved = [...heldContent(first.control), ...heldContent(second)];
  const joined =
    moved.length === 0
      ? first.control
      : withContent(first.control, moved, createNodeIdAllocator(part));
  const children = first.holder.children.flatMap((child) => {
    if (child.id === second.id) return [];
    return child.id === first.control.id ? [joined] : [child];
  });
  return written(
    part,
    first.holder,
    children,
    effectOf(first.paragraphId, [], [second.id]),
    options
  );
}

/** What a control holds as content: nothing while it shows its prompt. */
function heldContent(control: OoxmlElement): readonly OoxmlNode[] {
  return isShowingPlaceholder(control) ? [] : contentControlContentOf(control)!.children;
}

/** The control holding `content`, no longer showing its prompt. */
function withContent(
  control: OoxmlElement,
  content: readonly OoxmlNode[],
  nextId: () => string
): OoxmlElement {
  const sdtContent = contentControlContentOf(control)!;
  const properties = editedProperties(
    contentControlPropertiesContainerOf(control),
    { showingPlaceholder: false },
    nextId
  );
  const sdtPr = contentControlPropertiesContainerOf(control);
  return {
    ...control,
    children: control.children.map((child) => {
      if (child.id === sdtPr?.id) return properties;
      return child.id === sdtContent.id
        ? ({ ...sdtContent, children: content } as OoxmlNode)
        : child;
    }),
  } as OoxmlElement;
}

function withProperties(control: OoxmlElement, properties: OoxmlElement): OoxmlElement {
  const sdtPr = contentControlPropertiesContainerOf(control);
  return {
    ...control,
    children: sdtPr
      ? control.children.map((child) => (child.id === sdtPr.id ? properties : child))
      : [properties, ...control.children],
  } as OoxmlElement;
}

function effectOf(
  paragraphId: string,
  created: readonly string[],
  deleted: readonly string[]
): TreeOpEffect {
  return {
    dirty: [paragraphId],
    created,
    deleted,
    dependencyKeys: TEXT_DEPS,
    impact: 'paragraph-local',
  };
}

function written(
  part: OoxmlPart,
  holder: OoxmlElement,
  children: readonly OoxmlNode[],
  effect: TreeOpEffect,
  options?: EditOptions
): TreeOpResult {
  const edit = replaceChildren(part, holder.id, children, options);
  if (!edit.ok) return { ok: false, reason: 'tree-invariant', detail: JSON.stringify(edit.issues) };
  return { ok: true, part: edit.part, effect };
}
