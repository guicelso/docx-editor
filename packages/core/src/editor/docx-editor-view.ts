// The facade's view lane: Word's Print or Web Layout, and the width the web view lays out at.
//
// The twin of the zoom lane (`docx-editor-zoom.ts`). The web view's page is as wide as the
// scroller's content box, so a resize lays it out again, coalesced to one frame the way a fit
// is: a dragged window delivers a callback per frame, and each one would re-lay the document.

import type { EditorLayoutView } from '../contracts/editor-view.ts';
import type { ExecResult } from '../contracts/types.ts';
import type { PaginatedSurface, PaginatedSurfaceOptions } from './paginated-surface.ts';
import { scrollerContentWidth } from './scroller-content-width.ts';
import type { SurfaceLayoutView } from './surface-layout-view.ts';
import { surfaceScroller } from './surface-pages.ts';

/** What the lane reads from, and writes back to, the facade around it. */
export interface ViewLaneHost {
  container(): HTMLElement | null;
  /** Read every call: the surface is rebuilt on load and on the font remount. */
  surface(): PaginatedSurface | null;
  bump(): void;
  emitSelectionChange(): void;
}

export interface ViewLane {
  /** The view in force: the mounted surface's, or the requested one before a document mounts. */
  view(): EditorLayoutView;
  /** The view a new surface mounts in, when the scroller can be measured now. */
  mountOptions(): Pick<PaginatedSurfaceOptions, 'layoutView'>;
  /** Follow the scroller's width. Called after every mount. */
  attach(): void;
  detach(): void;
  readonly members: {
    getView(): EditorLayoutView;
    setView(view: EditorLayoutView): ExecResult;
  };
}

const PRINT_VIEW: SurfaceLayoutView = Object.freeze({ kind: 'print' });

export function createViewLane(
  requested: EditorLayoutView | undefined,
  host: ViewLaneHost
): ViewLane {
  let wanted: EditorLayoutView = requested ?? 'print';
  /**
   * The surface that has already answered the requested web view. A document it refused is
   * not asked again on every resize: only a new surface or a new request asks again.
   */
  let answered: PaginatedSurface | null = null;
  let offeredAtMount = false;
  let observer: ResizeObserver | null = null;
  let observed: HTMLElement | null = null;
  let frame: number | null = null;

  function layoutViewOf(view: EditorLayoutView): SurfaceLayoutView | null {
    if (view === 'print') return PRINT_VIEW;
    const container = host.container();
    const widthPx = container ? scrollerContentWidth(container) : null;
    return widthPx === null ? null : { kind: 'web', widthPx };
  }

  const viewInForce = (): EditorLayoutView => host.surface()?.state().layoutView ?? wanted;

  /** Lay out at the scroller's width now, publishing only when the view itself moved. */
  function follow(): void {
    const surface = host.surface();
    if (!surface || wanted !== 'web') return;
    if (answered === surface && surface.state().layoutView !== 'web') return;
    const next = layoutViewOf('web');
    if (next === null) return;
    const before = surface.state().layoutView;
    answered = surface;
    surface.setLayoutView(next);
    if (surface.state().layoutView === before) return;
    host.bump();
    host.emitSelectionChange();
  }

  function cancelFrame(): void {
    if (frame === null) return;
    observed?.ownerDocument.defaultView?.cancelAnimationFrame(frame);
    frame = null;
  }

  function schedule(): void {
    const view = observed?.ownerDocument.defaultView;
    if (!view?.requestAnimationFrame) return follow();
    if (frame !== null) return;
    frame = view.requestAnimationFrame(() => {
      frame = null;
      follow();
    });
  }

  function detach(): void {
    cancelFrame();
    observer?.disconnect();
    observer = null;
    observed = null;
  }

  function observe(): void {
    detach();
    const container = host.container();
    const scroller = container ? surfaceScroller(container) : null;
    if (wanted !== 'web' || !scroller || typeof ResizeObserver === 'undefined') return;
    observed = scroller;
    observer = new ResizeObserver(schedule);
    observer.observe(scroller);
  }

  function setView(next: EditorLayoutView): ExecResult {
    if (next !== 'print' && next !== 'web') {
      return { ok: false, code: 'invalidArgs', reason: `unknown view ${JSON.stringify(next)}` };
    }
    const before = viewInForce();
    const surface = host.surface();
    if (surface) {
      const layoutView = layoutViewOf(next);
      if (layoutView === null) {
        return {
          ok: false,
          code: 'unsupported',
          reason: 'the web layout view needs a scroll container with a width',
        };
      }
      answered = surface;
      if (!surface.setLayoutView(layoutView)) {
        return {
          ok: false,
          code: 'unsupported',
          reason: surface.state().lastRejection ?? 'the document refuses the web layout view',
        };
      }
    }
    wanted = next;
    observe();
    const changed = viewInForce() !== before;
    host.bump();
    host.emitSelectionChange();
    return { ok: true, changed };
  }

  return {
    view: viewInForce,
    mountOptions() {
      const layoutView = wanted === 'web' ? layoutViewOf('web') : null;
      offeredAtMount = layoutView !== null;
      return layoutView === null ? {} : { layoutView };
    },
    attach() {
      answered = offeredAtMount ? host.surface() : null;
      offeredAtMount = false;
      observe();
      follow();
    },
    detach,
    members: { getView: viewInForce, setView },
  };
}
