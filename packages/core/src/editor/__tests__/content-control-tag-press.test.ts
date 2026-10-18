// A press on a content control's view-only tag goes to the host when the host asks for it: the
// tag is found by the layout's hit test, in viewing too, and a claimed press moves no caret.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import type { ContentControlTagPress } from '../../contracts/editor-content-control-view.ts';
import type { DocxEditorInstance } from '../docx-editor-types.ts';
import { docx } from './paginated-surface-fixtures.ts';
import { mountAnchorEditor } from './scroll-to-anchor-fixture.ts';

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const control = (tag: string, inner: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w:richText/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

const FIRST = docx(`<w:p>${run('um ')}${control('a', run('A'))}${run(' fim')}</w:p>`);
const SECOND = docx(`<w:p>${run('dois ')}${control('b', run('B'))}${run(' fim')}</w:p>`);

const cleanups: (() => void)[] = [];
afterEach(() => {
  for (const cleanup of cleanups.splice(0).reverse()) cleanup();
});

function mounted(): { editor: DocxEditorInstance; host: HTMLElement } {
  const opened = mountAnchorEditor(FIRST);
  cleanups.push(opened.destroy);
  opened.editor.setContentControlTags({
    labelsOf: ({ tag }) => ({ open: { text: `${tag} ▸` }, close: { text: '◂' } }),
  });
  return opened;
}

/** The middle of the painted tag at one edge, in client coordinates. */
function tagPoint(editor: DocxEditorInstance, edge: 'open' | 'close') {
  const page = editor.surface!.layout().pages[0]!;
  const scale = editor.getRenderScale();
  for (const fragment of page.fragments) {
    if (fragment.kind !== 'paragraph') continue;
    for (const line of fragment.lines) {
      const span = line.spans.find((candidate) => candidate.contentControlTag?.edge === edge);
      if (!span) continue;
      return {
        clientX: (page.contentBox.x + span.box.x + span.box.width / 2) * scale,
        clientY: (page.contentBox.y + span.box.y + span.box.height / 2) * scale,
      };
    }
  }
  throw new Error(`no ${edge} tag painted`);
}

/** Press the middle of a tag, as a reader would. Returns the press, so its default can be read. */
function pressTag(
  editor: DocxEditorInstance,
  host: HTMLElement,
  edge: 'open' | 'close',
  init: PointerEventInit = {}
): PointerEvent {
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
  const point = tagPoint(editor, edge);
  const base = { bubbles: true, cancelable: true, button: 0, pointerId: 1, pointerType: 'mouse' };
  const press = new PointerEvent('pointerdown', { ...base, ...point, ...init });
  pages.dispatchEvent(press);
  document.dispatchEvent(new PointerEvent('pointerup', { ...base, ...point }));
  return press;
}

const caret = (editor: DocxEditorInstance) => {
  const { anchor, head } = editor.surface!.state().selection;
  return [anchor.offset, head.offset];
};

describe('a press on a content control tag', () => {
  test('reaches the host with the control and the edge, and a claimed press moves no caret', () => {
    const { editor, host } = mounted();
    const presses: ContentControlTagPress[] = [];
    editor.setContentControlTagPress((press) => {
      presses.push(press);
      return true;
    });
    const before = caret(editor);

    const event = pressTag(editor, host, 'open');

    expect(presses).toHaveLength(1);
    expect(presses[0]!.control.tag).toBe('a');
    expect(presses[0]!.edge).toBe('open');
    expect(event.defaultPrevented).toBe(true);
    expect(caret(editor)).toEqual(before);
  });

  test('a claimed press still reaches the outside-press listeners of the page', () => {
    const { editor, host } = mounted();
    editor.setContentControlTagPress(() => true);
    let reached = false;
    const listener = () => {
      reached = true;
    };
    document.addEventListener('pointerdown', listener);

    pressTag(editor, host, 'close');
    document.removeEventListener('pointerdown', listener);

    expect(reached).toBe(true);
  });

  test('reaches the host while the document is open for viewing', () => {
    const { editor, host } = mounted();
    const tags: (string | undefined)[] = [];
    editor.setContentControlTagPress((press) => {
      tags.push(press.control.tag);
      return true;
    });
    editor.setEditingMode('view');

    pressTag(editor, host, 'open');

    expect(tags).toEqual(['a']);
  });

  test('a press the host declines places the caret as before', () => {
    const { editor, host } = mounted();
    let asked = 0;
    editor.setContentControlTagPress(() => {
      asked += 1;
      return false;
    });

    const event = pressTag(editor, host, 'close');

    expect(asked).toBe(1);
    expect(event.defaultPrevented).toBe(true);
    const [anchor, head] = caret(editor);
    expect(anchor).toBe(head);
  });

  test('a press with a modifier is never asked', () => {
    const { editor, host } = mounted();
    let asked = 0;
    editor.setContentControlTagPress(() => {
      asked += 1;
      return true;
    });

    pressTag(editor, host, 'open', { shiftKey: true });

    expect(asked).toBe(0);
  });

  test('the policy reaches the next document, and null takes it away', () => {
    const { editor, host } = mounted();
    const tags: (string | undefined)[] = [];
    editor.setContentControlTagPress((press) => {
      tags.push(press.control.tag);
      return true;
    });
    editor.load(SECOND);

    pressTag(editor, host, 'open');
    editor.setContentControlTagPress(null);
    pressTag(editor, host, 'open');

    expect(tags).toEqual(['b']);
  });
});
