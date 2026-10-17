/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The web layout view and the host's write admission are local to one participant. A peer in
// the web view keeps its single page while the other edits on sheets; the admission refuses the
// local reader's write and never a remote one, so the replicas converge through concurrent
// typing, undo and a save/reopen.

import { afterEach, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { collaborationModule } from '../collaboration-module.ts';
import { createPeerHarness, zipDocument } from './document-peer-support.ts';
import { packageFingerprint, saveReopenDigest } from './document-support.ts';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const harness = createPeerHarness('web-view-admission', { offlineEditing: true });
type Editor = ReturnType<typeof createDocxEditor>;
const editors: Editor[] = [];
const scrollers: HTMLElement[] = [];
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
  for (const scroller of scrollers.splice(0)) scroller.remove();
  harness.cleanup();
});

/** An editor inside a measured scroller, the way a host's viewport holds it. */
function mount(options: Parameters<typeof createDocxEditor>[0]): Editor {
  const scroller = document.createElement('div');
  scroller.className = 'docx-editor__scroll-container';
  Object.defineProperty(scroller, 'clientWidth', { value: 480, configurable: true });
  Object.defineProperty(scroller, 'clientHeight', { value: 600, configurable: true });
  const container = document.createElement('div');
  scroller.append(container);
  document.body.append(scroller);
  scrollers.push(scroller);
  const editor = createDocxEditor({ ...options, container, zoomMode: { type: 'fixed' } });
  editors.push(editor);
  return editor;
}

function caretAtEnd(editor: Editor): void {
  const session = editor.surface!.session;
  const position = { paragraphId: session.paragraphIds()[0]!, offset: session.bodyText().length };
  editor.surface!.setSelection({ anchor: position, head: position });
}

const SENTENCE = '<w:p><w:r><w:t>abc</w:t></w:r></w:p>';

async function peers() {
  const pair = await harness.pair(zipDocument(SENTENCE));
  pair.alice.detach();
  pair.bob.detach();
  const alice = mount({
    document: pair.alice.room.document,
    view: 'web',
    modules: [collaborationModule({ session: pair.alice.room.session })],
  });
  const bob = mount({
    document: pair.bob.room.document,
    modules: [collaborationModule({ session: pair.bob.room.session })],
  });
  alice.setWriteAdmission(({ part }) =>
    JSON.stringify(part.root).includes('X') ? 'no X in a sentence' : null
  );
  const flush = () => {
    pair.alice.room.session.flushPendingJournals();
    pair.bob.room.session.flushPendingJournals();
  };
  const converge = () => {
    flush();
    expect(packageFingerprint(alice.surface!.session.currentPackage())).toBe(
      packageFingerprint(bob.surface!.session.currentPackage())
    );
  };
  return { pair, alice, bob, flush, converge };
}

test('the admission refuses the local write and admits the remote one', async () => {
  const { alice, bob, converge } = await peers();
  expect(alice.getView()).toBe('web');
  expect(bob.getView()).toBe('print');

  caretAtEnd(alice);
  alice.surface!.type('X');
  expect(alice.snapshot().lastRejection).toBe('no X in a sentence');
  converge();
  expect(bob.surface!.session.bodyText()).toBe('abc');

  caretAtEnd(bob);
  bob.surface!.type('X');
  converge();
  expect(alice.surface!.session.bodyText()).toBe('abcX');
  expect(alice.getView()).toBe('web');
  expect(alice.getTotalPages()).toBe(1);
});

test('concurrent admitted typing converges, through undo and a save/reopen', async () => {
  const { pair, alice, bob, flush, converge } = await peers();
  pair.pause();
  caretAtEnd(alice);
  alice.surface!.type('A');
  caretAtEnd(bob);
  bob.surface!.type('B');
  expect(alice.snapshot().lastRejection).toBeNull();
  flush();
  pair.resume();
  converge();
  const text = alice.surface!.session.bodyText();
  expect(text.slice(0, 3)).toBe('abc');
  expect(text.slice(3).split('').sort().join('')).toBe('AB');

  expect(alice.exec({ type: 'undo' }).ok).toBe(true);
  converge();
  expect(bob.surface!.session.bodyText()).toBe('abcB');

  const reopened = mount({ document: new Uint8Array(await alice.save()), view: 'web' });
  expect(saveReopenDigest(reopened.surface!.session.currentPackage())).toEqual(
    saveReopenDigest(alice.surface!.session.currentPackage())
  );
  expect(reopened.getView()).toBe('web');
});
