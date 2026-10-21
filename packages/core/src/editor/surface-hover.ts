// What a resting pointer is on: the one answer the published hover and the `data-hover` retint
// of a control's chrome both read, from the layout hit test a press uses.
//
// A control's value widget is painted outside it, past its right edge, so the control keeps the
// hover over the widget and over the gap to it: otherwise the button the pointer is reaching for
// fades on the way there. The widget's place is the painter's own rule, not a pixel slack.

import type { EditorHover, EditorHoverField } from '../contracts/editor.ts';
import {
  contentControlAtPoint,
  type HitPoint,
  type SemanticHit,
} from '../layout/semantic-hit-test.ts';
import { unionLayoutBoxes } from '../layout/semantic-record-queries.ts';
import type {
  ContentControlBoundaryRecord,
  LayoutBox,
  PageRecord,
  SemanticLayout,
} from '../layout/semantic-records.ts';
import { contentControlWidgetArea } from '../output/semantic-paint-content-controls.ts';

export const NO_HOVER: EditorHover = Object.freeze({ controlId: null, field: null, box: null });

/** A hover whose box is still in page-content coordinates, before the surface places it on screen. */
export interface PageHover {
  readonly controlId: string | null;
  readonly field: EditorHoverField | null;
  readonly box: LayoutBox | null;
}

/** Whether two hovers name the same pieces; where they are painted is not their identity. */
export function sameHoveredPiece(a: EditorHover, b: EditorHover): boolean {
  return a.controlId === b.controlId && sameField(a.field, b.field);
}

/**
 * The pieces under a point of a page's content area: the field only when the point is on its
 * glyphs, and the innermost control whose boundary — or, failing every boundary, whose widget
 * reach — holds the point. The box is the innermost piece's, where the point is.
 */
export function hoverOnPage(
  layout: SemanticLayout,
  page: PageRecord,
  point: HitPoint,
  hit: SemanticHit
): PageHover {
  const control =
    contentControlAtPoint(layout, page.index, point) ?? controlByWidgetReach(page, point);
  const field = hit.onGlyphs && hit.field ? hit.field : null;
  if (field) {
    const { paragraphId, start, end, instruction, box } = field;
    return { controlId: control?.id ?? null, field: { paragraphId, start, end, instruction }, box };
  }
  return {
    controlId: control?.id ?? null,
    field: null,
    box: control ? fragmentUnder(page, control, point) : null,
  };
}

function controlByWidgetReach(
  page: PageRecord,
  point: HitPoint
): ContentControlBoundaryRecord | null {
  return (
    (page.contentControls ?? []).find((control) => {
      const reach = widgetReachOf(page, control);
      return reach !== null && contains(reach, point);
    }) ?? null
  );
}

/** The control's first box on the page joined to its widget: the widget, and the gap before it. */
function widgetReachOf(page: PageRecord, control: ContentControlBoundaryRecord): LayoutBox | null {
  const widget = contentControlWidgetArea(page, control);
  const first = control.fragments.find((fragment) => fragment.pageIndex === page.index);
  return widget && first ? unionLayoutBoxes([first.box, widget]) : null;
}

/** The fragment of the control the point is in; the widget's reach lands on its first box. */
function fragmentUnder(
  page: PageRecord,
  control: ContentControlBoundaryRecord,
  point: HitPoint
): LayoutBox | null {
  const onPage = control.fragments.filter((fragment) => fragment.pageIndex === page.index);
  return (onPage.find((fragment) => contains(fragment.box, point)) ?? onPage[0])?.box ?? null;
}

function contains(box: LayoutBox, point: HitPoint): boolean {
  return (
    point.x >= box.x &&
    point.x < box.x + box.width &&
    point.y >= box.y &&
    point.y < box.y + box.height
  );
}

function sameField(a: EditorHoverField | null, b: EditorHoverField | null): boolean {
  if (a === null || b === null) return a === b;
  return (
    a.paragraphId === b.paragraphId &&
    a.start === b.start &&
    a.end === b.end &&
    a.instruction === b.instruction
  );
}
