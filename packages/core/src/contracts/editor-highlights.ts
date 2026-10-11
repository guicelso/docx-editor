import type { ViewScope } from './editor-scope.ts';

/**
 * One text range to mark. A `TextMatch` from `findMatches()` is a valid range, so search
 * results pass straight through. @public
 */
export interface HighlightRange {
  /** Paragraph that holds the range: `TextMatch.blockId`. IDs belong to the open document. */
  readonly blockId: string;
  /** UTF-16 offset in the paragraph's model text, as `TextMatch.start` counts it. */
  readonly start: number;
  /** Range length in the same units. A zero length marks nothing. */
  readonly length: number;
  /** Story that holds the paragraph. Informational; `blockId` already names the paragraph. */
  readonly scope?: ViewScope;
  /**
   * Model text the range must cover to paint. Set it for ranges computed outside the editor,
   * such as on a server, so a range that no longer matches the document never paints. When it
   * is omitted, a range must keep covering the text it covered when it was first set.
   */
  readonly expectedText?: string;
}

/**
 * A content control to mark whole: its content in every paragraph it reaches and the tags drawn
 * at its edges and at its children's — the area its boundary outlines, at either level. @public
 */
export interface HighlightControl {
  /** The control's node id, as `ContentControlBoundaryRecord.id` names it. */
  readonly controlId: string;
}

/**
 * A place between two characters to mark, such as where a pending insertion lands. It paints
 * as an insertion bar with an optional label above it, the way a collaborator's caret paints,
 * and it never takes room in the line. @public
 */
export interface HighlightPosition {
  /** Paragraph that holds the position: `TextMatch.blockId`. IDs belong to the open document. */
  readonly blockId: string;
  /**
   * UTF-16 offset in the paragraph's model text, from 0 to the paragraph's length. It counts in
   * the document as it is when the set is installed.
   */
  readonly offset: number;
  /** Plain text drawn in a flag above the bar. It reaches the page as text, never as markup. */
  readonly label?: string;
}

/** What one highlight marks: a text range, a content control whole, or a position. @public */
export type HighlightTarget = HighlightRange | HighlightControl | HighlightPosition;

/** Presentation for one named highlight set. Every field is optional. @public */
export interface HighlightOptions {
  /**
   * Fill for every range: a CSS color or `var()` expression.
   * Default: `var(--doc-text-highlight-color)`, a translucent yellow.
   */
  readonly color?: string;
  /**
   * Index of the active range, or `-1` for none. An index past the last range also means
   * none. The active range takes `activeColor` and the `docx-text-highlight--active` class.
   * Default: `-1`.
   */
  readonly activeIndex?: number;
  /** Fill for the active range. Default: `var(--doc-text-highlight-active-color)`, orange. */
  readonly activeColor?: string;
  /** Space-separated CSS classes added to every mark of this set. */
  readonly className?: string;
  /**
   * Stacking order between sets, as an integer from -1000 to 1000. A higher value paints
   * on top. Sets with equal priority stack in the order they were first set. Default: `0`.
   * The built-in Find pane uses the set name `search` with priority `10`.
   */
  readonly priority?: number;
  /**
   * How the marks composite over the page. `'tint'` multiplies over light paper and screens
   * over dark paper, so the glyphs under a mark keep their color — a highlighter. `'cover'`
   * paints the color as is, over every tint set, so an opaque or page-colored fill hides or
   * dims the text — a redaction or a focus veil. Default: `'tint'`.
   */
  readonly blend?: HighlightBlend;
}

/** How a highlight set composites over the page; see {@link HighlightOptions.blend}. @public */
export type HighlightBlend = 'tint' | 'cover';

/** What `setHighlights()` resolved against the open document. @public */
export interface HighlightResult {
  /** Ranges that resolve to their text in the open document and can paint. */
  readonly applied: number;
  /**
   * Targets that do not resolve or cannot paint: an unknown paragraph, offsets past the
   * paragraph end, a zero length, text that differs from `expectedText`, a control the
   * document does not hold, a position an edit replaced, or a story without a layout position.
   */
  readonly unavailable: number;
}

/** A client-space rectangle, compatible with `DOMRect` readers such as positioning libraries. @public */
export interface HighlightRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
  readonly left: number;
}

/** What every painted mark under a point reports, whatever it marks. @public */
export interface HighlightMarkHit<R extends HighlightTarget> {
  /** Set name passed to `setHighlights()`. */
  readonly name: string;
  /** Index of the target in the array passed to `setHighlights()`. */
  readonly index: number;
  /** The target object exactly as it was passed, including your own fields. */
  readonly range: R;
  /** Whether this target is the set's active one. */
  readonly active: boolean;
  /** The marked box under the point, in client coordinates. Anchor popovers to it. */
  readonly rect: HighlightRect;
}

/** A mark of a text range under a point. `range` keeps the offsets you passed. @public */
export interface HighlightRangeHit<
  R extends HighlightRange = HighlightRange,
> extends HighlightMarkHit<R> {
  /** Current start of the highlighted text in its paragraph, after edits moved it. */
  readonly start: number;
  /** Current length of the highlighted text. */
  readonly length: number;
}

/** A mark of a whole content control under a point. @public */
export interface HighlightControlHit<
  R extends HighlightControl = HighlightControl,
> extends HighlightMarkHit<R> {
  /** The content control the mark covers. */
  readonly controlId: string;
}

/** A mark of a position under a point: the bar, without its label. @public */
export interface HighlightPositionHit<
  R extends HighlightPosition = HighlightPosition,
> extends HighlightMarkHit<R> {
  /** Current offset of the position in its paragraph, after edits moved it. */
  readonly offset: number;
}

/**
 * One painted mark under a point. `R` is your target type: pass targets with extra fields,
 * such as a glossary definition, and read them back from `range`. @public
 */
export type HighlightHit<R extends HighlightTarget = HighlightRange> = R extends HighlightControl
  ? HighlightControlHit<R>
  : R extends HighlightPosition
    ? HighlightPositionHit<R>
    : R extends HighlightRange
      ? HighlightRangeHit<R>
      : never;

/**
 * Paint-only text highlights: search results, glossary terms, review findings.
 *
 * Highlights are view state. They never change the document, selection, focus, or undo
 * history, they are never saved, and they are not shared with collaborators. @public
 */
export interface EditorHighlights {
  /**
   * Mark text ranges, content controls and positions as one named set, replacing the set's
   * previous targets and options. A control target marks the control whole, tags included; it
   * stops painting when the document no longer holds the control. A position target paints as
   * is, above every tint and cover set, whatever `blend` says: it covers no text to tint.
   *
   * Names are 1 to 64 letters, digits, `-`, or `_`, and start with a letter. Each set paints
   * in its own layer, so sets never replace each other. An empty `ranges` array clears the
   * set. At most 32 sets exist at once. A set paints its first 10000 ranges and counts the
   * rest as unavailable. The name `search` belongs to the shared document search.
   *
   * Ranges cover the body, tables, headers, footers, footnotes, endnotes, and anchored text
   * boxes in the body. A range moves with its text when an edit before it shifts the
   * paragraph. When an edit changes the text inside a range, the range stops painting, so a
   * mark never covers the wrong text. Search again after edits to mark new occurrences. A
   * position moves the same way; an edit that replaces text around it removes it until the
   * host sets it again, because no offset then names the place it marked.
   *
   * Loading, refreshing, or recovering the document removes every set. Invalid names,
   * ranges, or options throw
   * `TypeError` or `RangeError` and leave the current highlights unchanged.
   *
   * @example
   * ```ts
   * const matches = editor.findMatches('Supplier', { wholeWord: true });
   * editor.setHighlights('search', matches, { activeIndex: 0 });
   * ```
   * @public
   */
  setHighlights(
    name: string,
    ranges: readonly HighlightTarget[],
    options?: HighlightOptions
  ): HighlightResult;
  /** Remove one highlight set, or every set when `name` is omitted. @public */
  clearHighlights(name?: string): void;
  /**
   * The marks under a client point, topmost first. Use it from `pointermove` or `click`
   * handlers on the editor to show definitions or details for a marked term.
   * Only marks on painted pages are reported. @public
   */
  getHighlightsAt<R extends HighlightTarget = HighlightRange>(
    clientX: number,
    clientY: number
  ): readonly HighlightHit<R>[];
}
