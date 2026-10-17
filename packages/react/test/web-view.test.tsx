// The adapter half of the web layout view: `DocxEditor.Root` hands `view` to the engine and
// applies a later change without a remount, and `DocxEditor.Viewport` says which view is in
// force so the stylesheet can drop the sheet's breathing room.

// MUST be first: happy-dom registration happens on import.
import './dom-setup.ts';

(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

import { afterEach, beforeEach, expect, test } from 'bun:test';
import { act, cleanup, render } from '@testing-library/react';
import { strToU8, zipSync } from 'fflate';
import type { EditorLayoutView } from '@docx-editor.dev/core/contracts/editor';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { DocxEditorRoot } from '../src/editor/DocxEditorRoot.tsx';
import { DocxEditorViewport } from '../src/editor/DocxEditorViewport.tsx';
import { DocxEditorContent } from '../src/editor/DocxEditorContent.tsx';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';

const SOURCE = zipSync({
  '[Content_Types].xml': strToU8(
    `<Types xmlns="${CT}">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '</Types>'
  ),
  '_rels/.rels': strToU8(
    `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
  ),
  'word/document.xml': strToU8(
    `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>hello world</w:t></w:r></w:p></w:body></w:document>`
  ),
});

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
  try {
    cleanup();
  } finally {
    (globalThis as { ResizeObserver?: unknown }).ResizeObserver = RealResizeObserver;
  }
});

function tree(view: EditorLayoutView, ready: (editor: DocxEditorInstance) => void) {
  return (
    <DocxEditorRoot
      document={SOURCE}
      view={view}
      zoomMode={{ type: 'fixed' }}
      onReady={(editor) => ready(editor as DocxEditorInstance)}
    >
      <DocxEditorViewport>
        <DocxEditorContent />
      </DocxEditorViewport>
    </DocxEditorRoot>
  );
}

test('Root opens in the web view at the viewport width and switches back without a remount', async () => {
  let instance: DocxEditorInstance | null = null;
  const rendered = render(tree('web', (editor) => (instance = editor)));
  const scroll = rendered.container.querySelector('[data-testid="docx-editor-scroll"]')!;
  Object.defineProperty(scroll, 'clientWidth', { value: 600, configurable: true });
  await act(async () => {
    for (const observer of [...observers]) observer();
    await new Promise((resolve) => requestAnimationFrame(resolve));
  });

  const editor = instance!;
  expect(editor.getView()).toBe('web');
  expect(scroll.getAttribute('data-docx-view')).toBe('web');
  expect(editor.getPageGeometry()[0]!.box.width).toBeCloseTo(600, 6);

  await act(async () => {
    rendered.rerender(tree('print', (next) => (instance = next)));
  });

  expect(instance).toBe(editor);
  expect(editor.getView()).toBe('print');
  expect(scroll.hasAttribute('data-docx-view')).toBe(false);
});
