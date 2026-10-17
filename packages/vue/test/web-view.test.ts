// The Vue twin of the React web view test: `DocxEditorRoot` hands `view` to the engine and
// applies a later change without a remount, and `DocxEditorViewport` says which view is in force.

import './dom-setup.ts';

import { afterEach, beforeEach, expect, test } from 'bun:test';
import { createApp, h, reactive } from 'vue';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { DocxEditorRoot } from '../src/editor/DocxEditorRoot';
import { DocxEditorViewport } from '../src/editor/DocxEditorViewport';
import { DocxEditorContent } from '../src/editor/DocxEditorContent';
import { flush, SOURCE } from './helpers/mount';

let observers: (() => void)[] = [];
const RealResizeObserver = globalThis.ResizeObserver;

beforeEach(() => {
  observers = [];
  (globalThis as { ResizeObserver?: unknown }).ResizeObserver = class {
    constructor(callback: (entries: unknown[]) => void) {
      observers.push(() => callback([]));
    }
    observe(): void {}
    unobserve(): void {}
    disconnect(): void {}
  };
});

afterEach(() => {
  (globalThis as { ResizeObserver?: unknown }).ResizeObserver = RealResizeObserver;
  document.body.innerHTML = '';
});

test('DocxEditorRoot opens in the web view at the viewport width and switches back in place', async () => {
  const props = reactive({ view: 'web' as 'web' | 'print' });
  const ready: DocxEditorInstance[] = [];
  const container = document.createElement('div');
  document.body.append(container);
  const app = createApp({
    render: () =>
      h(
        DocxEditorRoot,
        {
          document: SOURCE,
          view: props.view,
          zoomMode: { type: 'fixed' },
          onReady: (editor: unknown) => ready.push(editor as DocxEditorInstance),
        },
        { default: () => h(DocxEditorViewport, null, { default: () => h(DocxEditorContent) }) }
      ),
  });
  app.mount(container);
  try {
    await flush();
    const scroll = container.querySelector('[data-testid="docx-editor-scroll"]')!;
    Object.defineProperty(scroll, 'clientWidth', { value: 600, configurable: true });
    for (const observer of [...observers]) observer();
    await new Promise((resolve) => requestAnimationFrame(resolve));
    await flush();

    const editor = ready.at(-1)!;
    expect(editor.getView()).toBe('web');
    expect(scroll.getAttribute('data-docx-view')).toBe('web');
    expect(editor.getPageGeometry()[0]!.box.width).toBeCloseTo(600, 6);

    props.view = 'print';
    await flush();

    expect(ready.at(-1)).toBe(editor);
    expect(editor.getView()).toBe('print');
    expect(scroll.hasAttribute('data-docx-view')).toBe(false);
  } finally {
    app.unmount();
    container.remove();
  }
});
