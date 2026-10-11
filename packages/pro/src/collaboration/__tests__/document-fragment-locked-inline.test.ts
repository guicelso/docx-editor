/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A fragment that carries a content-locked control lands inside a line that already holds one.
// Two peers must converge on it, with every lock intact, through undo, redo and a save/reopen.

import { afterEach, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import {
  contentControlPropertiesOf,
  contentControlTextOf,
  contentControlsIn,
  readOoxmlPart,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { collaborationModule } from '../collaboration-module.ts';
import { createPeerHarness, zipDocument } from './document-peer-support.ts';
import { mainPart, packageFingerprint, saveReopenDigest } from './document-support.ts';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const harness = createPeerHarness('fragment-locked-inline', { offlineEditing: true });
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
  return editor;
}

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const run = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const locked = (tag: string, held: string): string =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w:lock w:val="contentLocked"/></w:sdtPr>` +
  `<w:sdtContent>${held}</w:sdtContent></w:sdt>`;

/** The fragment a host lands: a locked control that holds text and a locked citation. */
function carried(): OoxmlNode {
  const read = readOoxmlPart(
    `<w:p xmlns:w="${W}">${locked('entry', run('Pre: ') + locked('cited', run('ACME')))}</w:p>`,
    { name: '/fragment.xml', contentType: 'application/xml' }
  );
  if (!read.ok) throw new Error(read.reason);
  return read.part.root;
}

async function peers() {
  const pair = await harness.pair(
    zipDocument(`<w:p>${locked('first', run('IGLOO'))}${run(' e o vendedor')}</w:p>`)
  );
  const [alice, bob] = [pair.alice, pair.bob].map((peer) => {
    peer.detach();
    return mount({
      document: peer.room.document,
      modules: [collaborationModule({ session: peer.room.session })],
    });
  }) as [Editor, Editor];
  const converge = () => {
    pair.alice.room.session.flushPendingJournals();
    pair.bob.room.session.flushPendingJournals();
    expect(packageFingerprint(alice.surface!.session.currentPackage())).toBe(
      packageFingerprint(bob.surface!.session.currentPackage())
    );
  };
  return { alice, bob, converge };
}

/** Each control by tag, with its own text and lock. */
function controls(editor: Editor) {
  return contentControlsIn(mainPart(editor.surface!.session.currentPackage()).root).map(
    ({ node }) => ({
      tag: contentControlPropertiesOf(node).tag,
      lock: contentControlPropertiesOf(node).lock,
      text: contentControlTextOf(node),
    })
  );
}

test('a peer lands a locked control inside a line that holds one; both converge through undo and reopen', async () => {
  const { alice, bob, converge } = await peers();
  const paragraphId = alice.surface!.session.paragraphIds()[0]!;

  const written = alice.surface!.applyAutomationOps(() => [
    { op: 'insertFragment', paragraphId, offset: 10, blocks: [carried()] },
  ]);

  expect(written.committed).toBe(true);
  converge();
  expect(bob.surface!.session.bodyText()).toBe('IGLOO e o Pre: ACMEvendedor');
  expect(controls(bob)).toEqual([
    { tag: 'first', lock: 'contentLocked', text: 'IGLOO' },
    { tag: 'entry', lock: 'contentLocked', text: 'Pre: ACME' },
    { tag: 'cited', lock: 'contentLocked', text: 'ACME' },
  ]);

  expect(alice.exec({ type: 'undo' }).ok).toBe(true);
  converge();
  expect(bob.surface!.session.bodyText()).toBe('IGLOO e o vendedor');
  expect(alice.exec({ type: 'redo' }).ok).toBe(true);
  converge();
  expect(controls(bob).map(({ tag }) => tag)).toEqual(['first', 'entry', 'cited']);

  const reopened = mount({ document: new Uint8Array(await bob.save()) });
  expect(saveReopenDigest(reopened.surface!.session.currentPackage())).toEqual(
    saveReopenDigest(bob.surface!.session.currentPackage())
  );
  expect(controls(reopened)).toEqual(controls(bob));
});
