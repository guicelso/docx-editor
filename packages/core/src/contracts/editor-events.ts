import type { ResolvedRevisionMarkup } from './revision-markup.ts';
import type { EditorError, EditorSnapshot } from './editor.ts';
import type { HistoryDiagnostic } from './editor-scope.ts';
import type { FieldSubject } from './editor-content-control-view.ts';

/**
 * What `editor.on(...)` can be subscribed to, and what each handler receives.
 *
 * These are PUSH notifications and are not interchangeable with reading `snapshot()`: a snapshot
 * read cannot observe an event that was never emitted, which is why adapter behaviour is asserted
 * against these rather than against the snapshot.
 */
export interface EditorEvents {
  /** Local viewer preferences changed through the API or review dialog. */
  revisionMarkupChange: (settings: ResolvedRevisionMarkup) => void;
  /** A document mutation committed, with the ids it touched. */
  change: (change: DocumentChange) => void;
  /** The selection or its derived formatting moved. */
  selectionChange: (snapshot: EditorSnapshot) => void;
  error: (error: EditorError) => void;
  historyDiagnostic: (diagnostic: HistoryDiagnostic) => void;
  /**
   * The piece under a resting pointer changed. Its own channel: a hover moves no selection, and
   * carried by `selectionChange` it re-derived the snapshot for every subscriber on each move.
   */
  hoverChange: (hover: EditorHover) => void;
}

/**
 * What a resting pointer is on, from the layout hit test a press uses — never the painted DOM.
 * `field` and `controlId` can both be set: a field inside a control is the innermost piece, and
 * `box` is where that innermost piece is painted under the pointer. All null when nothing is.
 */
export interface EditorHover {
  /** Innermost content control under the pointer, its view-only tags and its value widget included. */
  readonly controlId: string | null;
  /** The field atom the pointer is on its glyphs of; never one merely beside the pointer. */
  readonly field: EditorHoverField | null;
  /** In viewport pixels, read when the hover changed: a host anchors its own chrome to it. */
  readonly box: EditorHoverBox | null;
}

/** A field atom a pointer rests on: the field, and its model range. */
export interface EditorHoverField extends FieldSubject {
  /** Canonical node id of the paragraph the field is in. */
  readonly paragraphId: string;
  /** UTF-16 offset where the field starts in its paragraph. */
  readonly start: number;
  /** UTF-16 offset where the field ends in its paragraph. */
  readonly end: number;
}

/** A rectangle in viewport pixels, the space of `getBoundingClientRect`. */
export interface EditorHoverBox {
  /** Left edge. */
  readonly x: number;
  /** Top edge. */
  readonly y: number;
  /** Horizontal extent. */
  readonly width: number;
  /** Vertical extent. */
  readonly height: number;
}

/**
 * The payload of the `change` event / `onChange`. It carries revision + identity
 * deltas, NOT serialized bytes: serializing a whole DOCX on every keystroke would
 * be prohibitive for large documents. Call `save()` to get bytes on demand.
 */
export interface DocumentChange {
  /** Replacement events must not start another external processing request. Absent for edits. */
  readonly source?: 'load' | 'refresh' | 'recovery';
  /** Package revision after this change — `getDocumentHandle()`'s number, monotonic. */
  readonly revision: number;
  /** Block ids created/deleted/edited by this change, when the engine reports them. */
  readonly created?: readonly string[];
  readonly deleted?: readonly string[];
  readonly dirty?: readonly string[];
}
