/**
 * What a host shows over content controls and fields: an answer per control or per field, never
 * written to the document, exported or printed.
 */

/**
 * The host's names for a piece the engine draws, for the host's own stylesheet. Each lands on the
 * piece's painted elements as `data-piece-tone`, `data-piece-form`, `data-piece-state` and
 * `data-piece-role`. Letters, digits and `-`, starting with a letter, at most 32; a name outside
 * that is not published. @public
 */
export interface PieceName {
  /** What colours the piece, as `data-piece-tone`. */
  readonly tone?: string;
  /** How the piece takes its place in the text, as `data-piece-form`. */
  readonly form?: string;
  /** A state the host knows of and the engine does not, as `data-piece-state`. */
  readonly state?: string;
  /**
   * The part the piece plays in the structure around it, as `data-piece-role`: what it is to its
   * neighbours, which no interaction changes.
   */
  readonly role?: string;
}

/**
 * A piece the engine lays out as one unit: the room a chip keeps on each side and, when there is a
 * mark, a measured slot before its text that the host's {@link MarkPainter} fills. The line never
 * breaks inside it, and the caret stops only at its two edges. @public
 */
export interface PieceAtom extends PieceName {
  /**
   * The host's name for the mark, published as `data-piece-mark`, with the same pattern as a name.
   * The host's `MarkPainter` draws in its slot.
   */
  readonly mark?: string;
}

/**
 * What a host draws at one edge of a content control: a label, an edge that draws nothing, or
 * labels one after the other. Each is a tag, with a caret slot on each side of it. @public
 */
export type ContentControlTagLabel =
  | ContentControlTagInk
  | ContentControlTagEdgeOnly
  | ContentControlTagSequence;

/** A label drawn at a control's edge, with the host's names for it. @public */
export interface ContentControlTagInk extends PieceName {
  /** What the tag shows, laid out as one unbreakable unit. */
  readonly text: string;
  /**
   * `chip` (the default) draws the label in a neutral pill with room on each side. `text` draws
   * it as text in the line, with no room around it and no fill: punctuation that reads as the
   * document's own, styled by its names.
   */
  readonly variant?: 'chip' | 'text';
  /** A chip's mark, measured before its text as a {@link PieceAtom}'s is. A `text` label has none. */
  readonly mark?: string;
}

/**
 * An edge that draws nothing and takes no room, and still stands as a tag: the place just inside
 * the control is its own slot, as it is beside a label. For a control a host draws no label at
 * on one side, whose content must still be reachable from that side. @public
 */
export interface ContentControlTagEdgeOnly {
  readonly variant: 'edge';
}

/**
 * Labels drawn one after the other at one edge, as ONE tag: they lay out as a single unbreakable
 * unit, each keeping the room its own variant keeps, and the edge still has exactly its two slots.
 * For an edge that says two things, such as a condition and the text the document prints there. A
 * label with no text is left out; a sequence left with none draws no tag. @public
 */
export interface ContentControlTagSequence {
  readonly variant: 'sequence';
  readonly labels: readonly ContentControlTagInk[];
}

/** A content control as the file states it. `tag` is untrusted file data. @public */
export interface ContentControlSubject {
  /** The control's node id, as `ContentControlBoundaryRecord.id` names it. */
  readonly controlId: string;
  /** The control's `w:tag`, or undefined when the file states none. */
  readonly tag: string | undefined;
}

/**
 * The host's start and end tags for every content control, or null to draw none (Word's Design
 * Mode). @public
 *
 * `labelsOf` must be pure and cheap: it runs for every control of every paragraph a layout pass
 * prepares. Install a new display whenever an answer may differ. The engine asks every control
 * again and lays out only the paragraphs whose tags changed.
 */
export interface ContentControlTagDisplay {
  /** The start and end tags of one control; a side left out, or null, draws no tag there. */
  readonly labelsOf: (
    control: ContentControlSubject
  ) => { readonly open?: ContentControlTagLabel; readonly close?: ContentControlTagLabel } | null;
}

/**
 * The text a content control shows while it holds its placeholder (`w:showingPlcHdr`), in place
 * of the placeholder text the file stores. @public
 *
 * `promptOf` must be pure and cheap, like {@link ContentControlTagDisplay.labelsOf}. The shown
 * text takes the style of the stored placeholder and is one unit: the caret stops before and
 * after it, and the document, the saved bytes and the clipboard keep the stored text.
 */
export interface ContentControlPromptDisplay {
  /**
   * The text one control shows, or null (or empty) to show the stored placeholder. Asked only
   * of a control showing its placeholder, whose placeholder is plain text: inline, or a block
   * control holding one paragraph.
   */
  readonly promptOf: (control: ContentControlSubject) => string | null;
}

/**
 * What a host makes of one content control. @public
 *
 * An `atom` is met as one unit, as a control showing its prompt is: its content, when it is text
 * in one face, lays out as one {@link PieceAtom}; the caret steps over it in one step, a press
 * selects it whole, a drag takes it whole, and a key with it selected removes the control. A `frame` keeps its content as
 * it is and publishes the names on the control's chrome. Either way the chrome carries the names.
 */
export type ContentControlPiece = { readonly atom: PieceAtom } | { readonly frame: PieceName };

/**
 * The host's piece for every content control. @public
 *
 * `pieceOf` must be pure and cheap, like {@link ContentControlTagDisplay.labelsOf}, and the engine
 * lays out only the paragraphs whose answers changed. A control showing its prompt shows the
 * prompt, whatever the answer.
 */
export interface ContentControlPieceDisplay {
  /** The piece one control is, or null for a control the host does not draw as one. */
  readonly pieceOf: (control: ContentControlSubject) => ContentControlPiece | null;
}

/** A press on a content control's view-only tag: the control the tag labels and its edge. @public */
export interface ContentControlTagPress {
  /** The control the pressed tag labels. */
  readonly control: ContentControlSubject;
  /** Which of its tags: the one at its start or the one at its end. */
  readonly edge: 'open' | 'close';
}

/**
 * Whether the host takes a press on a view-only tag, as a link it draws there. @public
 *
 * Asked for a plain primary press (no modifier) that the layout's hit test resolves to a tag, in
 * viewing too. True takes it: the press places no caret and makes no selection. False, or no
 * policy, leaves it to place the caret beside the tag. Either way the press stays visible to the
 * page's outside-press listeners, so a host closing something on a press elsewhere still sees it.
 */
export type ContentControlTagPressPolicy = (press: ContentControlTagPress) => boolean;

/**
 * The host's face for a field, from its instruction (untrusted file text), or undefined to leave
 * the field as the engine draws it. A field with a face is a {@link PieceAtom}. @public
 *
 * Pure and cheap, like {@link ContentControlTagDisplay.labelsOf}: the engine asks it for every
 * field of every paragraph a layout pass prepares, and lays out only the paragraphs whose answers
 * changed.
 */
export type FieldFace = (instruction: string) => PieceAtom | undefined;

/**
 * The host's view over content controls and fields. @public
 *
 * Each setter installs a policy, a function the engine asks per control or per field. The editor
 * keeps every policy across `load`, `attach` and `detach`, and hands it to each document it
 * mounts. `null` removes it. Nothing reaches the document, the saved bytes, the clipboard or a
 * collaborator.
 */
export interface EditorContentControlView {
  /** Draw a start and an end tag at each content control, as the host labels it. */
  setContentControlTags(display: ContentControlTagDisplay | null): void;
  /** Show the host's text in each content control that holds its placeholder. */
  setContentControlPrompts(display: ContentControlPromptDisplay | null): void;
  /** Make each content control an atom or a named frame, as the host answers. */
  setContentControlPieces(display: ContentControlPieceDisplay | null): void;
  /** Give each field the host's face: its names and its mark, laid out as one unit. */
  setFieldFaces(face: FieldFace | null): void;
  /**
   * Which fields a plain press selects whole, from their instruction. The selected field's
   * painted result carries `data-selected`. Without a rule, a press places the caret at the
   * nearer edge of a field.
   */
  setFieldSelection(select: ((instruction: string) => boolean) | null): void;
  /** Take presses on the view-only tags the host draws, as it decides per tag. */
  setContentControlTagPress(press: ContentControlTagPressPolicy | null): void;
}
