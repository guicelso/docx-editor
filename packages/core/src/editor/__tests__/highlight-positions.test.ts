import { afterEach, describe, expect, test } from 'bun:test';
import type { HighlightPosition } from '../../contracts/editor-highlights.ts';
import { docx } from './paginated-surface-fixtures.ts';
import { mountAnchorEditor } from './scroll-to-anchor-fixture.ts';

const cleanups: (() => void)[] = [];
afterEach(() => {
  for (const cleanup of cleanups.splice(0).reverse()) cleanup();
});

const p = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

function mount() {
  const mounted = mountAnchorEditor(docx(p('The buyer pays the price.') + p('Second paragraph.')));
  cleanups.push(mounted.destroy);
  const [first] = mounted.editor.findMatches('buyer');
  const blockId = first!.blockId;
  const bars = () => [
    ...mounted.host.querySelectorAll<HTMLElement>('.docx-text-highlight-position'),
  ];
  const at = (offset: number, label?: string): HighlightPosition => ({
    blockId,
    offset,
    ...(label === undefined ? {} : { label }),
  });
  const caret = (offset: number) => {
    const position = { paragraphId: blockId, offset };
    mounted.editor.exec({ type: 'setSelection', range: { anchor: position, head: position } });
  };
  return { ...mounted, blockId, bars, at, caret };
}

const leftOf = (element: HTMLElement | undefined) => Number.parseFloat(element!.style.left);

describe('position targets', () => {
  test('paint a bar with its label as text, as view state only', () => {
    const { editor, bars, at } = mount();
    const canUndo = editor.snapshot().canUndo;
    let events = 0;
    editor.on('change', () => events++);

    expect(editor.setHighlights('slot', [at(4, '<b>aqui</b> entra')])).toEqual({
      applied: 1,
      unavailable: 0,
    });

    const [bar] = bars();
    expect(bar!.dataset.highlightIndex).toBe('0');
    expect(Number.parseFloat(bar!.style.width)).toBe(2);
    expect(Number.parseFloat(bar!.style.height)).toBeGreaterThan(0);
    const label = bar!.querySelector('.docx-text-highlight-position-label')!;
    expect(label.textContent).toBe('<b>aqui</b> entra');
    expect(label.querySelector('b')).toBeNull();
    expect(editor.snapshot().canUndo).toBe(canUndo);
    expect(events).toBe(0);
  });

  test('a later offset paints further right, and an empty label draws no flag', () => {
    const { editor, bars, at } = mount();
    editor.setHighlights('slot', [at(0), at(10, '')]);
    const [start, later] = bars();
    expect(leftOf(later)).toBeGreaterThan(leftOf(start));
    expect(later!.childElementCount).toBe(0);
  });

  test('positions paint as is, above the tint and cover groups, whatever the set blends', () => {
    const { editor, host, at } = mount();
    const [match] = editor.findMatches('price');
    editor.setHighlights('slot', [match!, at(4, 'aqui')]);
    const layer = host.querySelector('.docx-text-highlight-overlay')!;
    const groups = [...layer.children] as HTMLElement[];
    expect(groups.map((group) => group.dataset.highlightBlend ?? 'positions')).toEqual([
      'tint',
      'cover',
      'positions',
    ]);
    expect(groups[0]!.querySelectorAll('.docx-text-highlight')).toHaveLength(1);
    expect(groups[2]!.querySelectorAll('.docx-text-highlight-position')).toHaveLength(1);
    expect(
      groups[2]!.querySelector<HTMLElement>('[data-highlight-set]')!.dataset.highlightSet
    ).toBe('slot');
  });

  test('moves with an edit before it, and stays when text is typed at it', () => {
    const { editor, bars, at, caret } = mount();
    editor.setHighlights('slot', [at(4, 'aqui')]);
    const before = leftOf(bars()[0]);

    caret(4);
    expect(editor.exec({ type: 'insertText', text: 'X' }).ok).toBe(true);
    expect(leftOf(bars()[0])).toBe(before);

    caret(0);
    expect(editor.exec({ type: 'insertText', text: 'YY ' }).ok).toBe(true);
    expect(leftOf(bars()[0])).toBeGreaterThan(before);

    caret(20);
    const shifted = leftOf(bars()[0]);
    expect(editor.exec({ type: 'insertText', text: 'Z' }).ok).toBe(true);
    expect(leftOf(bars()[0])).toBe(shifted);
  });

  test('a deletion that ends at it carries it back to where the deletion began', () => {
    const { editor, bars, at, blockId } = mount();
    editor.setHighlights('slot', [at(10)]);
    const from = { paragraphId: blockId, offset: 4 };
    const to = { paragraphId: blockId, offset: 10 };
    editor.exec({ type: 'setSelection', range: { anchor: from, head: to } });
    expect(editor.exec({ type: 'deleteText' }).ok).toBe(true);
    editor.setHighlights('reference', [at(4)]);
    const [moved, reference] = bars();
    expect(leftOf(moved)).toBe(leftOf(reference));
  });

  test('an edit that replaces the text around it removes it until the host sets it again', () => {
    const { editor, bars, at, blockId } = mount();
    editor.setHighlights('slot', [at(6, 'aqui')]);
    const from = { paragraphId: blockId, offset: 4 };
    const to = { paragraphId: blockId, offset: 9 };
    editor.exec({ type: 'setSelection', range: { anchor: from, head: to } });
    expect(editor.exec({ type: 'insertText', text: 'client' }).ok).toBe(true);
    expect(bars()).toHaveLength(0);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(bars()).toHaveLength(0);

    expect(editor.setHighlights('slot', [at(6, 'aqui')])).toEqual({ applied: 1, unavailable: 0 });
    expect(bars()).toHaveLength(1);
  });

  test('counts a position past the paragraph end, or in a paragraph the document lacks', () => {
    const { editor, bars, at } = mount();
    expect(
      editor.setHighlights('slot', [at(4), at(500), { blockId: 'missing#1', offset: 0 }])
    ).toEqual({ applied: 1, unavailable: 2 });
    expect(bars()).toHaveLength(1);
  });

  test('a hit reports the bar and where the position is now', () => {
    const { editor, host, bars, at, caret } = mount();
    const target = at(4, 'aqui');
    editor.setHighlights('slot', [target], { activeIndex: 0 });
    caret(0);
    editor.exec({ type: 'insertText', text: 'XX ' });
    host.querySelector<HTMLElement>('.docx-text-highlight-overlay')!.getBoundingClientRect = () =>
      ({ left: 0, top: 0 }) as DOMRect;
    const [bar] = bars();
    expect(bar!.classList.contains('docx-text-highlight-position--active')).toBe(true);

    const [hit] = editor.getHighlightsAt<HighlightPosition>(
      leftOf(bar) + 1,
      Number.parseFloat(bar!.style.top) + 1
    );
    expect(hit!.range).toBe(target);
    expect(hit!.offset).toBe(7);
    expect(hit!.rect.width).toBe(2);
  });

  test('refuses a position that says more than one thing, and keeps the set unchanged', () => {
    const { editor, bars, at, blockId } = mount();
    editor.setHighlights('slot', [at(4)]);
    for (const target of [
      { blockId, offset: -1 },
      { blockId, offset: 1.5 },
      { blockId: '', offset: 0 },
      { blockId, offset: 0, start: 0 },
      { blockId, offset: 0, label: 5 },
      { controlId: 'c#1', offset: 0 },
    ]) {
      expect(() => editor.setHighlights('slot', [target as never])).toThrow(TypeError);
    }
    expect(bars()).toHaveLength(1);
  });
});
