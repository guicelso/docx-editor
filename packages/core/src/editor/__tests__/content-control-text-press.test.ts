// A press on text where content-control edges meet keeps the caret with the text it pressed: the
// end of a control's last word is inside the control, the start of the word after it is outside.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import type { DocxEditorInstance } from '../docx-editor-types.ts';
import { docx } from './paginated-surface-fixtures.ts';
import { mountAnchorEditor } from './scroll-to-anchor-fixture.ts';

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const control = (tag: string, inner: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w:richText/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

const DOCUMENT = docx(`<w:p>${run('um ')}${control('a', run('ABC'))}${run(' fim')}</w:p>`);

const cleanups: (() => void)[] = [];
afterEach(() => {
  for (const cleanup of cleanups.splice(0).reverse()) cleanup();
});

function mounted(): { editor: DocxEditorInstance; host: HTMLElement } {
  const opened = mountAnchorEditor(DOCUMENT);
  cleanups.push(opened.destroy);
  opened.editor.setContentControlTags({
    labelsOf: ({ tag }) => ({ open: { text: `${tag} ▸` }, close: { text: '◂' } }),
  });
  return opened;
}

/** A point just inside the start of the text painted from `start`, or the end of the one painted to `end`. */
function textPoint(editor: DocxEditorInstance, start: number, end: number, edge: 'start' | 'end') {
  const page = editor.surface!.layout().pages[0]!;
  const scale = editor.getRenderScale();
  for (const fragment of page.fragments) {
    if (fragment.kind !== 'paragraph') continue;
    for (const line of fragment.lines) {
      const span = line.spans.find(
        (candidate) =>
          !candidate.contentControlTag &&
          (edge === 'start' ? candidate.range.start === start : candidate.range.end === end)
      );
      if (!span) continue;
      const x = edge === 'start' ? span.box.x + 0.5 : span.box.x + span.box.width - 0.5;
      return {
        clientX: (page.contentBox.x + x) * scale,
        clientY: (page.contentBox.y + span.box.y + span.box.height / 2) * scale,
      };
    }
  }
  throw new Error(`no text ${start}–${end} painted`);
}

function press(host: HTMLElement, point: { clientX: number; clientY: number }): void {
  const pages = host.querySelector<HTMLElement>('.docx-pages')!;
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
  const base = { bubbles: true, cancelable: true, button: 0, pointerId: 1, pointerType: 'mouse' };
  pages.dispatchEvent(new PointerEvent('pointerdown', { ...base, ...point }));
  document.dispatchEvent(new PointerEvent('pointerup', { ...base, ...point }));
}

const caretOf = (editor: DocxEditorInstance) => {
  const { selection, contentControls } = editor.surface!.state();
  return { offset: selection.head.offset, slot: contentControls.caretSlot?.placement ?? null };
};

const controlIdOf = (editor: DocxEditorInstance): string =>
  editor.surface!.layout().contentControls?.find((record) => record.tag === 'a')?.id ?? '';

describe('a press on text where control edges meet', () => {
  test('the end of the last word of a control keeps the caret inside the control', () => {
    const { editor, host } = mounted();

    press(host, textPoint(editor, 3, 6, 'end'));

    expect(caretOf(editor)).toEqual({ offset: 6, slot: { inside: controlIdOf(editor) } });
  });

  test('the start of the word after a control keeps the caret outside, touching that word', () => {
    const { editor, host } = mounted();

    press(host, textPoint(editor, 6, 10, 'start'));

    expect(caretOf(editor)).toEqual({
      offset: 6,
      slot: { beside: { controlId: controlIdOf(editor), side: 'after' } },
    });
  });

  test('the end of the word before a control stays in front of it, and the start of its first word inside', () => {
    const { editor, host } = mounted();

    press(host, textPoint(editor, 0, 3, 'end'));

    expect(caretOf(editor)).toEqual({
      offset: 3,
      slot: { beside: { controlId: controlIdOf(editor), side: 'before' } },
    });

    press(host, textPoint(editor, 3, 6, 'start'));

    expect(caretOf(editor)).toEqual({ offset: 3, slot: { inside: controlIdOf(editor) } });
  });
});
