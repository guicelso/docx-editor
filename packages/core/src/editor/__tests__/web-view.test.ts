// The editor's web layout view and the host's write admission, through the facade: the page
// follows the viewport's content width, a document the view does not lay out is refused with
// the reason, a write that would bring such content in is refused the same way, and the host's
// admission speaks for every write and survives a load.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { paragraphTextOf } from '@docx-editor.dev/core/store';
import type { DocxEditorConfig, DocxEditorInstance } from '../docx-editor-types.ts';
import { docx, putCaret } from './paginated-surface-fixtures.ts';
import { createDocxEditor } from '../docx-editor.ts';

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const LONG = docx(
  Array.from({ length: 90 }, (_, index) => `<w:p>${run(`Line ${index}`)}</w:p>`).join('')
);
const COLUMNS = docx(`<w:p>${run('two columns')}</w:p><w:sectPr><w:cols w:num="2"/></w:sectPr>`);

const cleanups: (() => void)[] = [];
afterEach(() => {
  for (const cleanup of cleanups.splice(0).reverse()) cleanup();
});

/** An editor in a scroller `widthPx` wide, whose resizes the test delivers by hand. */
function mounted(bytes: Uint8Array, config: Partial<DocxEditorConfig>, widthPx = 480) {
  const observers: (() => void)[] = [];
  const realObserver = globalThis.ResizeObserver;
  globalThis.ResizeObserver = class {
    constructor(callback: () => void) {
      observers.push(callback);
    }
    observe(): void {}
    unobserve(): void {}
    disconnect(): void {}
  } as unknown as typeof ResizeObserver;
  let width = widthPx;
  const scroller = document.createElement('div');
  scroller.className = 'docx-editor__scroll-container';
  Object.defineProperty(scroller, 'clientWidth', { get: () => width, configurable: true });
  Object.defineProperty(scroller, 'clientHeight', { value: 600, configurable: true });
  const host = document.createElement('div');
  scroller.append(host);
  document.body.append(scroller);
  const editor = createDocxEditor({
    document: bytes,
    container: host,
    zoomMode: { type: 'fixed' },
    ...config,
  });
  cleanups.push(() => {
    editor.destroy();
    scroller.remove();
    globalThis.ResizeObserver = realObserver;
  });
  return {
    editor,
    async resize(next: number): Promise<void> {
      width = next;
      for (const observer of observers) observer();
      await new Promise((resolve) => setTimeout(resolve, 50));
    },
  };
}

/** The first page's width in CSS pixels at 100%. */
const pageWidthPx = (editor: DocxEditorInstance): number => editor.getPageGeometry()[0]!.box.width;

describe('the web layout view', () => {
  test('with no width to measure yet, waits on pages and takes the width when it comes', async () => {
    const opened = mounted(LONG, { view: 'web' }, 0);
    expect(opened.editor.getView()).toBe('print');

    await opened.resize(320);

    expect(opened.editor.getView()).toBe('web');
    expect(pageWidthPx(opened.editor)).toBeCloseTo(320, 6);
  });

  test('switches without remounting, and the page follows the scroller', async () => {
    const opened = mounted(LONG, {});
    const { editor } = opened;
    const surface = editor.surface;
    expect(editor.getView()).toBe('print');
    expect(editor.getTotalPages()).toBeGreaterThan(1);

    expect(editor.setView('web')).toEqual({ ok: true, changed: true });

    expect(editor.surface).toBe(surface);
    expect(editor.snapshot().view).toBe('web');
    expect(editor.getTotalPages()).toBe(1);
    expect(pageWidthPx(editor)).toBeCloseTo(480, 6);
    await opened.resize(360);
    expect(pageWidthPx(editor)).toBeCloseTo(360, 6);
    expect(editor.setView('print')).toEqual({ ok: true, changed: true });
    expect(editor.getTotalPages()).toBeGreaterThan(1);
  });

  test('opens in the web view when the document allows it', () => {
    const { editor } = mounted(LONG, { view: 'web' });

    expect(editor.getView()).toBe('web');
    expect(editor.getTotalPages()).toBe(1);
  });

  test('a document the web view does not lay out is refused, and stays on pages', () => {
    const { editor } = mounted(COLUMNS, {});

    const result = editor.setView('web');

    expect(result).toEqual({
      ok: false,
      code: 'unsupported',
      reason: 'the web layout view does not show columns',
    });
    expect(editor.getView()).toBe('print');
    expect(editor.snapshot().lastRejection).toBe('the web layout view does not show columns');
  });

  test('in the web view, a note cannot be written in', () => {
    const { editor } = mounted(LONG, { view: 'web' });
    putCaret(editor.surface!, 2);

    editor.surface!.insertNote('footnote');

    expect(editor.snapshot().lastRejection).toBe(
      'the web layout view does not show footnotes or endnotes'
    );
    expect(editor.surface!.session.part().root.children.length).toBeGreaterThan(0);
    expect(JSON.stringify(editor.surface!.session.part().root)).not.toContain('noteReference');
  });
});

describe('the write admission', () => {
  test('refuses a typed write in the host’s words, and is kept across a load', () => {
    const { editor } = mounted(docx(`<w:p>${run('abc')}</w:p>`), {});
    editor.setWriteAdmission(({ part }) =>
      JSON.stringify(part.root).includes('X') ? 'no X in a sentence' : null
    );
    putCaret(editor.surface!, 3);

    editor.surface!.type('X');

    expect(editor.snapshot().lastRejection).toBe('no X in a sentence');
    const surface = editor.surface!;
    expect(paragraphTextOf(surface.session.part(), surface.session.paragraphIds()[0]!)).toBe('abc');

    editor.load(docx(`<w:p>${run('def')}</w:p>`));
    putCaret(editor.surface!, 3);
    editor.surface!.type('X');

    expect(editor.snapshot().lastRejection).toBe('no X in a sentence');
    editor.setWriteAdmission(null);
    editor.surface!.type('Y');
    const reloaded = editor.surface!;
    expect(paragraphTextOf(reloaded.session.part(), reloaded.session.paragraphIds()[0]!)).toBe(
      'defY'
    );
  });
});
