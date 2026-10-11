/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A position highlight is view state of one peer: it never reaches the other peer or the saved
// bytes, and it stays on its place while the other peer edits the same paragraph.

import { afterEach, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { collaborationModule } from '../collaboration-module.ts';
import { createPeerHarness, zipDocument } from './document-peer-support.ts';
import { packageFingerprint, saveReopenDigest } from './document-support.ts';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const harness = createPeerHarness('highlight-positions', { offlineEditing: true });
type Editor = ReturnType<typeof createDocxEditor>;
const editors: Editor[] = [];
const containers: HTMLElement[] = [];
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
  for (const container of containers.splice(0)) container.remove();
  harness.cleanup();
});

function mount(options: Parameters<typeof createDocxEditor>[0]) {
  const container = document.createElement('div');
  document.body.append(container);
  containers.push(container);
  const editor = createDocxEditor({ ...options, container });
  editors.push(editor);
  return { editor, container };
}

const BODY = '<w:p><w:r><w:t xml:space="preserve">The buyer pays the price.</w:t></w:r></w:p>';

async function peers() {
  const pair = await harness.pair(zipDocument(BODY));
  const [alice, bob] = [pair.alice, pair.bob].map((peer) => {
    peer.detach();
    return mount({
      document: peer.room.document,
      modules: [collaborationModule({ session: peer.room.session })],
    });
  }) as [ReturnType<typeof mount>, ReturnType<typeof mount>];
  const converge = () => {
    pair.alice.room.session.flushPendingJournals();
    pair.bob.room.session.flushPendingJournals();
    expect(packageFingerprint(alice.editor.surface!.session.currentPackage())).toBe(
      packageFingerprint(bob.editor.surface!.session.currentPackage())
    );
  };
  return { alice, bob, converge };
}

const nextTask = (): Promise<void> => new Promise((resolve) => setTimeout(resolve, 0));

/** A remote edit lays out on the scheduler's own task, and paint may follow it by one more. */
async function settled(editor: Editor): Promise<void> {
  const surface = editor.surface!;
  for (let task = 0; task < 20; task += 1) {
    if (surface.publishedLayout().revision === surface.session.packageRevision()) break;
    await nextTask();
  }
  await nextTask();
}

const paragraphOf = (editor: Editor) => editor.surface!.session.paragraphIds()[0]!;

function select(editor: Editor, from: number, to = from) {
  const paragraphId = paragraphOf(editor);
  editor.surface!.setSelection({
    anchor: { paragraphId, offset: from },
    head: { paragraphId, offset: to },
  });
}

const bars = (container: HTMLElement) => [
  ...container.querySelectorAll<HTMLElement>('.docx-text-highlight-position'),
];

/** The offset the painted bar names now, read back through the hit test. */
function offsetOf(editor: Editor, container: HTMLElement): number | null {
  const [bar] = bars(container);
  if (!bar) return null;
  container.querySelector<HTMLElement>('.docx-text-highlight-overlay')!.getBoundingClientRect =
    () => ({ left: 0, top: 0 }) as DOMRect;
  const [hit] = editor.getHighlightsAt<{ blockId: string; offset: number }>(
    Number.parseFloat(bar.style.left) + 1,
    Number.parseFloat(bar.style.top) + 1
  );
  return hit && 'offset' in hit ? hit.offset : null;
}

test('a position follows a peer typing before it, and stays out of the shared document', async () => {
  const { alice, bob, converge } = await peers();
  expect(
    alice.editor.setHighlights('slot', [
      { blockId: paragraphOf(alice.editor), offset: 4, label: 'aqui entra o trecho' },
    ])
  ).toEqual({ applied: 1, unavailable: 0 });

  select(bob.editor, 0);
  bob.editor.surface!.type('Then ');
  converge();
  await settled(alice.editor);

  expect(alice.editor.surface!.session.bodyText()).toBe('Then The buyer pays the price.');
  expect(offsetOf(alice.editor, alice.container)).toBe(9);
  expect(bars(bob.container)).toHaveLength(0);

  select(bob.editor, 20);
  bob.editor.surface!.type('!');
  converge();
  await settled(alice.editor);
  expect(offsetOf(alice.editor, alice.container)).toBe(9);

  const reopened = mount({ document: new Uint8Array(await alice.editor.save()) }).editor;
  expect(saveReopenDigest(reopened.surface!.session.currentPackage())).toEqual(
    saveReopenDigest(bob.editor.surface!.session.currentPackage())
  );
});

test('a peer replacing the text around a position removes it on the other peer', async () => {
  const { alice, bob, converge } = await peers();
  alice.editor.setHighlights('slot', [{ blockId: paragraphOf(alice.editor), offset: 6 }]);
  expect(bars(alice.container)).toHaveLength(1);

  select(bob.editor, 4, 9);
  bob.editor.surface!.type('client');
  converge();
  await settled(alice.editor);

  expect(alice.editor.surface!.session.bodyText()).toBe('The client pays the price.');
  expect(bars(alice.container)).toHaveLength(0);
  expect(
    alice.editor.setHighlights('slot', [{ blockId: paragraphOf(alice.editor), offset: 6 }])
  ).toEqual({ applied: 1, unavailable: 0 });
});
