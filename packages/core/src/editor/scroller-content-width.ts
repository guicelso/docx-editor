// The room the page has: the scroller's content box, which the zoom fit and the web view read.

import { surfaceScroller } from './surface-pages.ts';

/** The scroller's content box, or null when it cannot be measured. */
export function scrollerContentWidth(container: HTMLElement): number | null {
  const scroller = surfaceScroller(container);
  if (!scroller) return null;
  const width = scroller.clientWidth;
  if (!Number.isFinite(width) || width <= 0) return null;
  const style = scroller.ownerDocument.defaultView?.getComputedStyle(scroller);
  if (!style) return width;
  // PHYSICAL, not logical. `clientWidth` is content + padding in physical terms, so these are
  // the two that reduce it whichever way the text runs — and the chrome above uses both
  // spellings (the comments rail sets `padding-right`, the navigation pane sets
  // `padding-inline-start`), which computed style resolves into these regardless.
  //
  // NaN from a value that is not px (a stylesheet that has not applied, a DOM that reports '')
  // means "no reservation known", not "give up" — hence `|| 0` rather than bailing out, which
  // would leave the fit permanently stale.
  const left = Number.parseFloat(style.paddingLeft) || 0;
  const right = Number.parseFloat(style.paddingRight) || 0;
  return Math.max(width - left - right, 0);
}
