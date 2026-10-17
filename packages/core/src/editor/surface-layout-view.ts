// The surface's view, sheets or Word's Web Layout, and the write admission it puts before the
// host's: a document in the web view must stay one the web view lays out.

import type { OoxmlPart, WriteAdmission } from '@docx-editor.dev/core/store';
import type { StyleCascadeTable } from '../layout/style-cascade.ts';
import {
  webLayoutViewRefusal,
  type LayoutView,
  type WebLayoutViewRefusal,
} from '../layout/web-layout-view.ts';

/** Sheets, or Web Layout at the host's content width in CSS pixels. */
export type SurfaceLayoutView =
  | { readonly kind: 'print' }
  | { readonly kind: 'web'; readonly widthPx: number };

const PRINT_VIEW: SurfaceLayoutView = Object.freeze({ kind: 'print' });

const REFUSED_CONTENT: Readonly<Record<WebLayoutViewRefusal, string>> = {
  columns: 'columns',
  'line-grid': 'a document grid',
  notes: 'footnotes or endnotes',
  'floating-drawing': 'floating pictures or shapes',
  'floating-table': 'floating tables',
  'text-frame': 'text frames',
};

export interface SurfaceLayoutViewDeps {
  readonly container: HTMLElement;
  readonly bodyPart: () => OoxmlPart;
  readonly styles: () => StyleCascadeTable | undefined;
  /** Points to CSS pixels, read per layout: a zoom re-breaks the text at the same width. */
  readonly scale: () => number;
  readonly hostAdmission: WriteAdmission | undefined;
}

export function createSurfaceLayoutView(deps: SurfaceLayoutViewDeps) {
  let view = PRINT_VIEW;
  let hostAdmission = deps.hostAdmission;

  const widthPtOf = (widthPx: number): number => widthPx / deps.scale();

  function refusalOf(next: SurfaceLayoutView, body: OoxmlPart): string | null {
    if (next.kind === 'print') return null;
    if (!(next.widthPx > 0) || !Number.isFinite(next.widthPx)) {
      return 'the web layout view needs a positive width';
    }
    const refusal = webLayoutViewRefusal(body, deps.styles(), widthPtOf(next.widthPx));
    return refusal === null
      ? null
      : `the web layout view does not show ${REFUSED_CONTENT[refusal]}`;
  }

  const admission: WriteAdmission = ({ part }) =>
    (part.name === deps.bodyPart().name ? refusalOf(view, part) : null) ??
    hostAdmission?.({ part }) ??
    null;

  return {
    kind: (): SurfaceLayoutView['kind'] => view.kind,
    layoutView: (): LayoutView | undefined =>
      view.kind === 'web' ? { kind: 'web', widthPt: widthPtOf(view.widthPx) } : undefined,
    isCurrent: (next: SurfaceLayoutView): boolean =>
      next.kind === view.kind && (next.kind === 'print' || next.widthPx === widthPxOf(view)),
    /** Take `next`, or say why the document refuses it and keep the view in force. */
    apply(next: SurfaceLayoutView): string | null {
      const refusal = refusalOf(next, deps.bodyPart());
      if (refusal !== null) return refusal;
      view = next;
      deps.container.classList.toggle('docx-paginated-surface--web', next.kind === 'web');
      return null;
    },
    admission,
    setHostAdmission(next: WriteAdmission | null): void {
      hostAdmission = next ?? undefined;
    },
  };
}

function widthPxOf(view: SurfaceLayoutView): number | null {
  return view.kind === 'web' ? view.widthPx : null;
}
