// The host's policy over structural edits: every paragraph break and deletion asks it first, from
// the keyboard and from an input method alike. `handled` leaves the document to the host,
// `default` lets the engine write what it always writes, and the editor keeps the policy across a
// load, as it keeps the host's view over controls.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { paragraphTextOf } from '@docx-editor.dev/core/store';
import type { EditIntent } from '../../contracts/editor-edit-policy.ts';
import type { DocxEditorInstance } from '../docx-editor-types.ts';
import { docx, putCaret } from './paginated-surface-fixtures.ts';
import { mountAnchorEditor } from './scroll-to-anchor-fixture.ts';

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const ONE = docx(`<w:p>${run('abcd')}</w:p>`);
const TWO = docx(`<w:p>${run('wxyz')}</w:p>`);

const cleanups: (() => void)[] = [];
afterEach(() => {
  for (const cleanup of cleanups.splice(0).reverse()) cleanup();
});

const hosts = new WeakMap<DocxEditorInstance, HTMLElement>();

function mounted(): DocxEditorInstance {
  const opened = mountAnchorEditor(ONE);
  cleanups.push(opened.destroy);
  hosts.set(opened.editor, opened.host);
  return opened.editor;
}

/** The painted pages the browser delivers keys and input to; looked up again after a load. */
function pages(editor: DocxEditorInstance): HTMLElement {
  return hosts.get(editor)!.querySelector<HTMLElement>('.docx-pages')!;
}

function key(editor: DocxEditorInstance, init: KeyboardEventInit): void {
  pages(editor).dispatchEvent(
    new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init })
  );
}

function input(editor: DocxEditorInstance, inputType: string): void {
  pages(editor).dispatchEvent(
    new InputEvent('beforeinput', { bubbles: true, cancelable: true, inputType })
  );
}

function texts(editor: DocxEditorInstance): string[] {
  const surface = editor.surface!;
  const part = surface.session.part();
  return surface.session.paragraphIds().map((id) => paragraphTextOf(part, id) ?? '');
}

/** A policy that records every intent and answers as told. */
function recording(answer: 'handled' | 'default') {
  const asked: EditIntent[] = [];
  return {
    asked,
    policy: (intent: EditIntent) => {
      asked.push(intent);
      return answer;
    },
  };
}

describe('the edit policy', () => {
  test('a handled paragraph break writes nothing, by key and by input method', () => {
    const editor = mounted();
    const host = recording('handled');
    editor.setEditPolicy(host.policy);
    putCaret(editor.surface!, 2);

    key(editor, { key: 'Enter' });
    input(editor, 'insertParagraph');

    expect(host.asked).toEqual([{ kind: 'paragraphBreak' }, { kind: 'paragraphBreak' }]);
    expect(texts(editor)).toEqual(['abcd']);
  });

  test('a handled deletion writes nothing, and each says its direction and unit', () => {
    const editor = mounted();
    const host = recording('handled');
    editor.setEditPolicy(host.policy);
    putCaret(editor.surface!, 2);

    key(editor, { key: 'Backspace' });
    key(editor, { key: 'Delete', ctrlKey: true });
    input(editor, 'deleteContentForward');
    input(editor, 'deleteWordBackward');

    expect(host.asked).toEqual([
      { kind: 'delete', direction: 'backward', unit: 'character' },
      { kind: 'delete', direction: 'forward', unit: 'word' },
      { kind: 'delete', direction: 'forward', unit: 'character' },
      { kind: 'delete', direction: 'backward', unit: 'word' },
    ]);
    expect(texts(editor)).toEqual(['abcd']);
  });

  test('the default lets the engine write the edit', () => {
    const editor = mounted();
    editor.setEditPolicy(recording('default').policy);
    putCaret(editor.surface!, 2);

    key(editor, { key: 'Backspace' });
    key(editor, { key: 'Enter' });

    expect(texts(editor)).toEqual(['a', 'cd']);
  });

  test('a line break and a page break are not asked', () => {
    const editor = mounted();
    const host = recording('handled');
    editor.setEditPolicy(host.policy);
    putCaret(editor.surface!, 2);

    key(editor, { key: 'Enter', shiftKey: true });
    key(editor, { key: 'Enter', ctrlKey: true });

    expect(host.asked).toEqual([]);
  });

  test('the policy survives a load, and null removes it', () => {
    const editor = mounted();
    const host = recording('handled');
    editor.setEditPolicy(host.policy);
    editor.load(TWO);
    putCaret(editor.surface!, 2);

    key(editor, { key: 'Enter' });
    expect(host.asked).toEqual([{ kind: 'paragraphBreak' }]);
    expect(texts(editor)).toEqual(['wxyz']);

    editor.setEditPolicy(null);
    key(editor, { key: 'Enter' });
    expect(texts(editor)).toEqual(['wx', 'yz']);
  });
});
