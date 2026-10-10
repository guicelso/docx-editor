/**
 * What a host shows over content controls and fields: an answer per control or per field, never
 * written to the document, exported or printed.
 */

/**
 * What a host draws at one edge of a content control: a label, or an edge that draws nothing.
 * Both are a tag, with a caret slot on each side of it. @public
 */
export type ContentControlTagLabel = ContentControlTagInk | ContentControlTagEdgeOnly;

/** A label drawn at a control's edge. @public */
export interface ContentControlTagInk {
  /** What the tag shows, laid out as one unbreakable unit. */
  readonly text: string;
  /**
   * An opaque name the chip publishes as `data-cc-tag-tone`, so the host's stylesheet colours it
   * with its own theme. Letters, digits and `-`, starting with a letter, at most 32; anything else
   * draws the neutral chip.
   */
  readonly tone?: string;
  /**
   * `chip` (the default) draws the label in a neutral pill with room on each side. `text` draws
   * it as text in the line, with no room around it and no fill: punctuation that reads as the
   * document's own, styled by the tone.
   */
  readonly variant?: 'chip' | 'text';
}

/**
 * An edge that draws nothing and takes no room, and still stands as a tag: the place just inside
 * the control is its own slot, as it is beside a label. For a control a host draws no label at
 * on one side, whose content must still be reachable from that side. @public
 */
export interface ContentControlTagEdgeOnly {
  readonly variant: 'edge';
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
 * The host's name for a field, from its instruction (untrusted file text), or undefined for
 * none. Letters, digits and `-`, starting with a letter, at most 32: it lands in an attribute.
 * @public
 */
export type FieldTone = (instruction: string) => string | undefined;

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
  /**
   * Name each field for the host's stylesheet: the name lands on the field's painted result as
   * `data-field-tone`. Changing it repaints without laying anything out.
   */
  setFieldTones(tone: FieldTone | null): void;
  /**
   * Which fields a plain press selects whole, from their instruction. The selected field's
   * painted result carries `data-selected`. Without a rule, a press places the caret at the
   * nearer edge of a field.
   */
  setFieldSelection(select: ((instruction: string) => boolean) | null): void;
}
