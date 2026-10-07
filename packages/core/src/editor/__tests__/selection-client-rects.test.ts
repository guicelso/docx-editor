// Where the selection is on screen, from layout: what a host anchors a floating menu to.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import type { PaginatedSurface } from '../paginated-surface.ts';
import { mount, putCaret } from './paginated-surface-fixtures.ts';

const mounted: PaginatedSurface[] = [];
afterEach(() => {
  for (const surface of mounted.splice(0)) surface.destroy();
});

function opened(): PaginatedSurface {
  const { surface } = mount('<w:p><w:r><w:t>Selected words here</w:t></w:r></w:p>');
  mounted.push(surface);
  return surface;
}

describe('selection client rects', () => {
  test('a range is one rectangle per line, as wide as the selected text', () => {
    const surface = opened();
    const paragraphId = surface.session.paragraphIds()[0]!;
    surface.setSelection({ anchor: { paragraphId, offset: 0 }, head: { paragraphId, offset: 8 } });
    const rects = surface.selectionClientRects();
    expect(rects).toHaveLength(1);
    expect(rects[0]!.width).toBeGreaterThan(0);
    expect(rects[0]!.right).toBeCloseTo(rects[0]!.left + rects[0]!.width, 6);
  });

  test('a wider range is wider on screen', () => {
    const surface = opened();
    const paragraphId = surface.session.paragraphIds()[0]!;
    const widthOf = (end: number) => {
      surface.setSelection({
        anchor: { paragraphId, offset: 0 },
        head: { paragraphId, offset: end },
      });
      return surface.selectionClientRects()[0]!.width;
    };
    expect(widthOf(14)).toBeGreaterThan(widthOf(4));
  });

  test('a collapsed caret is a zero-width rectangle with the line height', () => {
    const surface = opened();
    putCaret(surface, 3);
    const [rect] = surface.selectionClientRects();
    expect(rect?.width).toBe(0);
    expect(rect?.height).toBeGreaterThan(0);
  });
});
