// A frame the host takes whole: a drag crossing its edge extends over it while one inside it stays
// free, a selection from one edge of it to the other is the selected piece — the outermost of those
// that share both edges — and the state publishes it.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import type { ContentControlView } from '../../layout/content-control-view.ts';
import type { SemanticPosition } from '../../layout/semantic-interaction.ts';
import {
  contentControlPropertiesOf,
  contentControlsIn,
} from '../../store/package/content-control-nodes.ts';
import type { PaginatedSurface } from '../paginated-surface.ts';
import { absorbUnitControls } from '../surface-pointer-units.ts';
import { mount } from './paginated-surface-fixtures.ts';

const text = (value: string) =>
  `<w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">${value}</w:t></w:r>`;
const inline = (tag: string, inner: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w:richText/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
const LIST = inline(
  'list',
  inline('item-1', text('cd')) + text('; ') + inline('item-2', text('ef'))
);

/** `ab⟦list ⟦item-1 cd⟧; ⟦item-2 ef⟧⟧gh`: the list runs 2–8, its items 2–4 and 6–8. */
const SENTENCE = `<w:p>${text('ab')}${LIST}${text('gh')}</w:p>`;

/** The list as the only content of an optional: both run 2–8. */
const NESTED = `<w:p>${text('ab')}${inline('optional', LIST)}${text('gh')}</w:p>`;

const paragraph = (value: string) => `<w:p>${text(value)}</w:p>`;
const block = (tag: string, inner: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w:richText/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

/** `a | block[b | c] | d`. */
const BLOCKS = paragraph('a') + block('block', paragraph('b') + paragraph('c')) + paragraph('d');

const WHOLE: ContentControlView = { pieces: { pieceOf: () => ({ frame: {}, whole: true }) } };
const FRAMED: ContentControlView = { pieces: { pieceOf: () => ({ frame: {} }) } };

const mounted: PaginatedSurface[] = [];
afterEach(() => {
  for (const surface of mounted.splice(0)) surface.destroy();
});

function open(
  body: string,
  view: ContentControlView
): { surface: PaginatedSurface; container: HTMLElement } {
  const opened = mount(body);
  mounted.push(opened.surface);
  opened.surface.setContentControlView(view);
  return opened;
}

function controlIdOf(surface: PaginatedSurface, tag: string): string {
  const found = contentControlsIn(surface.session.part().root).find(
    (entry) => contentControlPropertiesOf(entry.node).tag === tag
  );
  if (!found) throw new Error(`no control tagged ${tag}`);
  return found.node.id;
}

function paragraphIds(surface: PaginatedSurface): readonly string[] {
  return surface
    .layout()
    .pages[0]!.fragments.flatMap((fragment) =>
      fragment.kind === 'paragraph' ? [fragment.paragraphId] : []
    );
}

function positionIn(surface: PaginatedSurface, paragraphIndex: number) {
  const paragraphId = paragraphIds(surface)[paragraphIndex]!;
  return (offset: number): SemanticPosition => ({ paragraphId, offset });
}

const offsetsOf = (surface: PaginatedSurface): [number, number] => {
  const { anchor, head } = surface.state().selection;
  return [anchor.offset, head.offset];
};

/** The client point of an offset in the first paragraph, from the spans layout published. */
function pointAt(surface: PaginatedSurface, offset: number): { clientX: number; clientY: number } {
  const page = surface.layout().pages[0]!;
  for (const fragment of page.fragments) {
    if (fragment.kind !== 'paragraph') continue;
    for (const line of fragment.lines) {
      for (const span of line.spans) {
        const { start, end } = span.range;
        if (offset < start || offset > end || end === start) continue;
        const x = span.box.x + ((offset - start) / (end - start)) * span.box.width;
        return {
          clientX: page.contentBox.x + x,
          clientY: page.contentBox.y + line.box.y + line.box.height / 2,
        };
      }
    }
  }
  throw new Error(`no span at offset ${offset}`);
}

function pagesOf(container: HTMLElement): HTMLElement {
  const pages = container.querySelector<HTMLElement>('.docx-pages')!;
  Object.defineProperty(pages, 'getBoundingClientRect', {
    configurable: true,
    value: () => ({
      left: 0,
      top: 0,
      right: 2000,
      bottom: 4000,
      width: 2000,
      height: 4000,
      x: 0,
      y: 0,
    }),
  });
  return pages;
}

function pointer(
  type: string,
  point: { clientX: number; clientY: number },
  init: PointerEventInit = {}
) {
  return new PointerEvent(type, {
    bubbles: true,
    cancelable: true,
    button: 0,
    pointerId: 1,
    pointerType: 'mouse',
    ...point,
    ...init,
  });
}

/** A press at one offset dragged to another, in the first paragraph. */
function drag(
  opened: { surface: PaginatedSurface; container: HTMLElement },
  from: number,
  to: number
): void {
  const pages = pagesOf(opened.container);
  pages.dispatchEvent(pointer('pointerdown', pointAt(opened.surface, from)));
  document.dispatchEvent(pointer('pointermove', pointAt(opened.surface, to)));
  document.dispatchEvent(pointer('pointerup', pointAt(opened.surface, to)));
}

function absorbed(surface: PaginatedSurface, from: number, to: number): [number, number] {
  const at = positionIn(surface, 0);
  const { anchor, head } = absorbUnitControls(surface.layout(), { anchor: at(from), head: at(to) });
  return [anchor.offset, head.offset];
}

describe('a drag over a frame taken whole', () => {
  test('a drag crossing the list’s edge extends over the list', () => {
    const opened = open(SENTENCE, WHOLE);
    drag(opened, 1, 3);
    expect(offsetsOf(opened.surface)).toEqual([1, 8]);
  });

  test('a Shift+click crossing the edge extends the same way', () => {
    const opened = open(SENTENCE, WHOLE);
    const pages = pagesOf(opened.container);
    opened.surface.setSelection({
      anchor: positionIn(opened.surface, 0)(9),
      head: positionIn(opened.surface, 0)(9),
    });
    pages.dispatchEvent(pointer('pointerdown', pointAt(opened.surface, 7), { shiftKey: true }));
    document.dispatchEvent(pointer('pointerup', pointAt(opened.surface, 7), { shiftKey: true }));
    expect(offsetsOf(opened.surface)).toEqual([9, 2]);
  });

  test('a selection inside an item stays free, and one across two items takes both whole', () => {
    const opened = open(SENTENCE, WHOLE);
    expect(absorbed(opened.surface, 2, 3)).toEqual([2, 3]);
    expect(absorbed(opened.surface, 3, 7)).toEqual([2, 8]);
  });

  test('the orientation of the drag is kept', () => {
    const opened = open(SENTENCE, WHOLE);
    expect(absorbed(opened.surface, 9, 3)).toEqual([9, 2]);
  });

  test('a frame without the mark is text to the drag', () => {
    const opened = open(SENTENCE, FRAMED);
    drag(opened, 1, 3);
    expect(offsetsOf(opened.surface)).toEqual([1, 3]);
  });

  test('a drag from a paragraph before a block frame into it extends to the block’s end', () => {
    const opened = open(BLOCKS, WHOLE);
    const before = positionIn(opened.surface, 0);
    const inside = positionIn(opened.surface, 1);
    const last = positionIn(opened.surface, 2);
    const taken = absorbUnitControls(opened.surface.layout(), {
      anchor: before(0),
      head: inside(1),
    });
    expect(taken.head).toEqual(last(1));
  });
});

describe('the selected piece', () => {
  test('a selection covering the list exactly selects it, with the chrome’s highlight', () => {
    const { surface, container } = open(SENTENCE, WHOLE);
    const at = positionIn(surface, 0);
    surface.setSelection({ anchor: at(2), head: at(8) });
    const listId = controlIdOf(surface, 'list');
    expect(surface.state().contentControls.selectedControlId).toBe(listId);
    const chrome = [...container.querySelectorAll<HTMLElement>('[data-docx-content-control]')].find(
      (each) => each.dataset.selected !== undefined
    );
    expect(chrome).toBeDefined();
  });

  test('the next key removes the selected list with what it holds', () => {
    const { surface } = open(SENTENCE, WHOLE);
    const at = positionIn(surface, 0);
    surface.setSelection({ anchor: at(8), head: at(2) });
    surface.deleteBackward();
    const xml = strFromU8(unzipSync(surface.save())['word/document.xml']!);
    expect(xml).not.toContain('w:val="list"');
    expect(xml).not.toContain('cd');
    expect(xml).toContain('gh');
  });

  test('covering an optional that holds only the list selects the optional, the outermost', () => {
    const { surface } = open(NESTED, WHOLE);
    const at = positionIn(surface, 0);
    surface.setSelection({ anchor: at(2), head: at(8) });
    expect(surface.state().contentControls.selectedControlId).toBe(
      controlIdOf(surface, 'optional')
    );
  });

  test('covering one item exactly selects the item, inside its list', () => {
    const { surface } = open(SENTENCE, WHOLE);
    const at = positionIn(surface, 0);
    surface.setSelection({ anchor: at(6), head: at(8) });
    expect(surface.state().contentControls.selectedControlId).toBe(controlIdOf(surface, 'item-2'));
  });

  test('a block frame covered from its first paragraph’s start to its last’s end is selected', () => {
    const { surface } = open(BLOCKS, WHOLE);
    surface.setSelection({ anchor: positionIn(surface, 1)(0), head: positionIn(surface, 2)(1) });
    expect(surface.state().contentControls.selectedControlId).toBe(controlIdOf(surface, 'block'));
  });

  test('without the mark, or with a partial selection, nothing is selected', () => {
    const framed = open(SENTENCE, FRAMED).surface;
    framed.setSelection({ anchor: positionIn(framed, 0)(2), head: positionIn(framed, 0)(8) });
    expect(framed.state().contentControls.selectedControlId).toBeNull();
    const whole = open(SENTENCE, WHOLE).surface;
    whole.setSelection({ anchor: positionIn(whole, 0)(2), head: positionIn(whole, 0)(7) });
    expect(whole.state().contentControls.selectedControlId).toBeNull();
    expect(whole.state().contentControls.activeControlId).not.toBeNull();
  });

  test('Backspace beside the list’s tag selects the list, even inside an optional of its size', () => {
    const { surface } = open(NESTED, {
      ...WHOLE,
      tags: { labelsOf: ({ tag }) => ({ open: { text: `${tag}▸` }, close: { text: `◂${tag}` } }) },
    });
    const listId = controlIdOf(surface, 'list');
    const at = positionIn(surface, 0)(8);
    surface.setSelection(
      { anchor: at, head: at },
      { controlId: listId, edge: 'close', side: 'after' }
    );
    surface.deleteBackward();
    expect(surface.state().contentControls.selectedControlId).toBe(listId);
  });
});
