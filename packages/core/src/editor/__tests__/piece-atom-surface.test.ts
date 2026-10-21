// An atom on a mounted surface: a press selects it whole and the next key removes it, the caret
// steps over it in one step, and the atom itself carries its hover and its selection.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import type { ContentControlView } from '../../layout/content-control-view.ts';
import type { StyleSpanRecord } from '../../layout/semantic-records.ts';
import type { PaginatedSurface } from '../paginated-surface.ts';
import { mount } from './paginated-surface-fixtures.ts';

const text = (value: string) =>
  `<w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">${value}</w:t></w:r>`;
const BODY =
  `<w:p>${text('A ')}<w:sdt><w:sdtPr><w:tag w:val="entry:q1"/><w:lock w:val="contentLocked"/></w:sdtPr>` +
  `<w:sdtContent>${text('Qualificação')}</w:sdtContent></w:sdt>${text(' fim')}</w:p>`;

const ATOM: ContentControlView = {
  pieces: {
    pieceOf: ({ tag }) =>
      tag === 'entry:q1'
        ? { atom: { tone: 'citation', form: 'atom', mark: 'qualification' } }
        : null,
  },
};

const mounted: PaginatedSurface[] = [];
afterEach(() => {
  for (const surface of mounted.splice(0)) surface.destroy();
});

function surfaceWithAtom(): { surface: PaginatedSurface; container: HTMLElement } {
  const opened = mount(BODY);
  mounted.push(opened.surface);
  opened.surface.setContentControlView(ATOM);
  return opened;
}

/** The atom's span and the line it is on, as layout published them. */
function atomOf(surface: PaginatedSurface): {
  span: StyleSpanRecord;
  lineY: number;
  lineHeight: number;
} {
  for (const fragment of surface.layout().pages[0]!.fragments) {
    if (fragment.kind !== 'paragraph') continue;
    for (const line of fragment.lines) {
      const span = line.spans.find((candidate) => candidate.pieceAtom);
      if (span) return { span, lineY: line.box.y, lineHeight: line.box.height };
    }
  }
  throw new Error('no atom laid out');
}

/** A primary press on the middle of the atom, as a reader would make it. */
function pressAtom(surface: PaginatedSurface, container: HTMLElement): void {
  const pages = container.querySelector<HTMLElement>('.docx-pages')!;
  // happy-dom lays nothing out, so the one rectangle the pointer reads is stated.
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
  const page = surface.layout().pages[0]!;
  const { span, lineY, lineHeight } = atomOf(surface);
  const scale = 1;
  const init = {
    bubbles: true,
    cancelable: true,
    button: 0,
    pointerId: 1,
    pointerType: 'mouse',
    clientX: (page.contentBox.x + span.box.x + span.box.width / 2) * scale,
    clientY: (page.contentBox.y + lineY + lineHeight / 2) * scale,
  };
  pages.dispatchEvent(new PointerEvent('pointerdown', init));
  document.dispatchEvent(new PointerEvent('pointerup', init));
}

const atomElement = (container: HTMLElement) =>
  container.querySelector<HTMLElement>('[data-piece-control]');

describe('an atom control on the surface', () => {
  test('paints as one piece with the host’s names, its control, and its mark slot', () => {
    const { container } = surfaceWithAtom();
    const atom = atomElement(container)!;
    expect(atom.textContent).toBe('Qualificação');
    expect(atom.dataset.pieceTone).toBe('citation');
    expect(atom.dataset.pieceForm).toBe('atom');
    expect(atom.querySelector<HTMLElement>('[data-piece-mark]')?.dataset.pieceMark).toBe(
      'qualification'
    );
  });

  test('a press selects it whole: the atom and its chrome say so', () => {
    const { surface, container } = surfaceWithAtom();
    pressAtom(surface, container);
    const { anchor, head } = surface.state().selection;
    expect([anchor.offset, head.offset]).toEqual([2, 14]);
    expect(atomElement(container)?.getAttribute('data-selected')).toBe('true');
    const chrome = container.querySelector<HTMLElement>('[data-docx-content-control]');
    expect(chrome?.dataset.selected).toBe('');
    expect(chrome?.dataset.pieceForm).toBe('atom');
  });

  test('with it selected whole, Backspace removes the control and what it holds', () => {
    const { surface, container } = surfaceWithAtom();
    pressAtom(surface, container);
    surface.deleteBackward();
    const xml = strFromU8(unzipSync(surface.save())['word/document.xml']!);
    expect(xml).not.toContain('entry:q1');
    expect(xml).not.toContain('Qualificação');
    expect(xml).toContain('fim');
  });

  test('the caret steps over it in one step, and extending takes it whole', () => {
    const { surface } = surfaceWithAtom();
    const paragraphId = surface.state().selection.head.paragraphId;
    const at = (offset: number) => ({ paragraphId, offset });
    surface.setSelection({ anchor: at(2), head: at(2) });
    surface.navigate('right');
    expect(surface.state().selection.head.offset).toBe(14);
    surface.navigate('left');
    expect(surface.state().selection.head.offset).toBe(2);
    surface.navigate('right', true);
    const { anchor, head } = surface.state().selection;
    expect([anchor.offset, head.offset]).toEqual([2, 14]);
  });

  test('the atom under the pointer carries the hover, and loses it when the pointer leaves', () => {
    const { surface, container } = surfaceWithAtom();
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
    const page = surface.layout().pages[0]!;
    const { span, lineY, lineHeight } = atomOf(surface);
    const scale = 1;
    pages.dispatchEvent(
      new PointerEvent('pointermove', {
        bubbles: true,
        clientX: (page.contentBox.x + span.box.x + span.box.width / 2) * scale,
        clientY: (page.contentBox.y + lineY + lineHeight / 2) * scale,
      })
    );
    expect(atomElement(container)?.hasAttribute('data-hover')).toBe(true);
    pages.dispatchEvent(new PointerEvent('pointerleave', { bubbles: true }));
    expect(atomElement(container)?.hasAttribute('data-hover')).toBe(false);
  });
});

describe('field faces and the mark painter on the surface', () => {
  const FIELD =
    '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> MERGEFIELD "field:a" </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>«cpf»</w:t></w:r>' +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';
  const FACES: ContentControlView = {
    fieldFaces: (instruction) =>
      instruction.includes('"field:')
        ? { tone: 'variable', form: 'atom', mark: 'variable' }
        : undefined,
  };

  test('a face makes the field an atom with its names; without it the field is as it was', () => {
    const opened = mount(FIELD);
    mounted.push(opened.surface);
    const atom = () => opened.container.querySelector<HTMLElement>('[data-field-atom]')!;
    opened.surface.setContentControlView(FACES);
    expect(atom().dataset.pieceTone).toBe('variable');
    expect(atom().querySelector('[data-piece-mark]')).not.toBeNull();
    opened.surface.setContentControlView(null);
    expect(atom().dataset.pieceTone).toBeUndefined();
    expect(atom().querySelector('[data-piece-mark]')).toBeNull();
  });

  test('the mark painter fills every slot, and changing it lays nothing out', () => {
    const opened = mount(FIELD);
    mounted.push(opened.surface);
    opened.surface.setContentControlView(FACES);
    const laidOut = opened.surface.layout();
    opened.surface.setMarkPainter((mark, slot) => {
      slot.dataset.painted = mark;
    });
    expect(opened.surface.layout()).toBe(laidOut);
    const slot = () => opened.container.querySelector<HTMLElement>('[data-piece-mark]')!;
    expect(slot().dataset.painted).toBe('variable');
    expect(slot().getAttribute('aria-hidden')).toBe('true');
    opened.surface.setMarkPainter(null);
    expect(slot().dataset.painted).toBeUndefined();
  });
});
