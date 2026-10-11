// The position targets of a highlight set: a place between two characters, painted as an
// insertion bar with an optional label above it, the way a collaborator's caret paints.
//
// A position takes no room in the line. Layout never sees it: the bar's geometry is the
// caret's (`caretAt`), read from the published layout like every other highlight.

import type { HighlightPosition, HighlightTarget } from '../contracts/editor-highlights.ts';
import type { KeyedParagraphRect } from '../layout/paragraph-range-rects.ts';
import { storyContentOffset } from '../layout/selection-rects.ts';
import { caretAt } from '../layout/semantic-interaction.ts';
import type { SemanticLayout, TextMeasurer } from '../layout/semantic-records.ts';

/** The painted width of the bar, in CSS pixels, centred on the position. */
export const POSITION_BAR_PX = 2;

/** A validated position target. An empty label draws no flag. */
export interface PositionTarget {
  readonly blockId: string;
  readonly offset: number;
  readonly label: string | null;
}

/**
 * The position a target names, or null for a range or a control. A target that names a
 * position and a range or a control says two things about one mark, and is refused.
 */
export function positionTargetOf(target: HighlightTarget, index: number): PositionTarget | null {
  if (typeof target !== 'object' || target === null || !('offset' in target)) return null;
  const { blockId, offset, label } = target as HighlightPosition;
  const valid =
    typeof blockId === 'string' &&
    blockId.length > 0 &&
    Number.isSafeInteger(offset) &&
    offset >= 0 &&
    offset <= 0x7fffffff &&
    (label === undefined || typeof label === 'string') &&
    !('start' in target) &&
    !('length' in target) &&
    !('controlId' in target);
  if (!valid) {
    throw new TypeError(
      `ranges[${index}] names a position by a blockId string, a nonnegative integer offset, ` +
        'and an optional label string, and nothing else.'
    );
  }
  return { blockId, offset, label: label === undefined || label === '' ? null : label };
}

/** One live position to paint: its index in the set, and where it stands now. */
export interface LivePosition {
  readonly key: number;
  readonly blockId: string;
  readonly offset: number;
}

/**
 * The bar of each position on the named pages, in page-content coordinates, with no width:
 * the painter centres a {@link POSITION_BAR_PX} bar on it. A header, footer or note position
 * takes its story box's offset, as a collaborator's caret does.
 */
export function positionRects(
  layout: SemanticLayout,
  positions: readonly LivePosition[],
  pages?: ReadonlySet<number>,
  measurer?: TextMeasurer
): KeyedParagraphRect[] {
  const rects: KeyedParagraphRect[] = [];
  for (const { key, blockId, offset } of positions) {
    const caret = caretAt(layout, { paragraphId: blockId, offset }, measurer);
    if (!caret || (pages && !pages.has(caret.pageIndex))) continue;
    const shift = storyContentOffset(layout, blockId, caret.pageIndex);
    rects.push({
      key,
      pageIndex: caret.pageIndex,
      x: caret.x + shift.x,
      y: caret.y + shift.y,
      width: 0,
      height: caret.height,
    });
  }
  return rects;
}

/** What a pooled position element currently shows, so a repaint writes only what changed. */
const positionState = new WeakMap<HTMLElement, { key: string }>();

/**
 * Reuse the sheet's position element at `at`, writing only what changed. The label reaches
 * the page through `textContent` only: it is the host's text, never markup.
 */
export function writePositionMark(
  sheet: HTMLElement,
  at: number,
  mark: {
    readonly index: number;
    readonly active: boolean;
    readonly left: number;
    readonly top: number;
    readonly height: number;
  },
  label: string | null,
  classes: readonly string[]
): void {
  let element = sheet.children[at] as HTMLElement | undefined;
  if (!element) {
    element = sheet.ownerDocument.createElement('div');
    element.style.position = 'absolute';
    sheet.append(element);
  }
  const key = `${mark.index}|${mark.active ? 1 : 0}|${mark.left}|${mark.top}|${mark.height}|${label ?? ''}`;
  if (positionState.get(element)?.key === key) return;
  element.className = mark.active
    ? 'docx-text-highlight-position docx-text-highlight-position--active'
    : 'docx-text-highlight-position';
  if (classes.length > 0) element.classList.add(...classes);
  element.setAttribute('data-highlight-index', String(mark.index));
  element.style.left = `${mark.left}px`;
  element.style.top = `${mark.top}px`;
  element.style.width = `${POSITION_BAR_PX}px`;
  element.style.height = `${mark.height}px`;
  writeLabel(element, label);
  positionState.set(element, { key });
}

function writeLabel(element: HTMLElement, label: string | null): void {
  if (label === null) {
    element.replaceChildren();
    return;
  }
  let flag = element.firstElementChild as HTMLElement | null;
  if (!flag) {
    flag = element.ownerDocument.createElement('span');
    flag.className = 'docx-text-highlight-position-label';
    element.append(flag);
  }
  flag.textContent = label;
}
