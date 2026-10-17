// The host's write admission: asked at the commit, over each part a local write changes, as the
// part would be published. A refusal leaves the document, its revision and its history as they
// were, and the host's own words reach the caller.

import { describe, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import { openTreeSession, type TreeDocxSession } from '../../binding/tree-session.ts';
import { paragraphTextOf } from '../store/tree-ops.ts';
import { ORIGIN_IDS } from '../registry/frozen-ids.ts';
import type { OoxmlPart } from '../package/ooxml-tree.ts';
import type { WriteAdmission } from '../store/write-admission.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';

const BYTES = zipSync({
  '[Content_Types].xml': strToU8(
    `<Types xmlns="${CT}">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '</Types>'
  ),
  '_rels/.rels': strToU8(
    `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
  ),
  'word/document.xml': strToU8(
    `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body><w:p><w:r><w:t>abc</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
  ),
});

const BODY = { kind: 'body' } as const;

/** A session whose admission is whatever `current` holds when a write is asked about. */
function opened(): { session: TreeDocxSession; admit(next: WriteAdmission | undefined): void } {
  let current: WriteAdmission | undefined;
  const result = openTreeSession(BYTES, { writeAdmission: () => current });
  if (!result.ok) throw new Error(result.reason);
  return {
    session: result.session,
    admit: (next) => {
      current = next;
    },
  };
}

const textOf = (session: TreeDocxSession): string =>
  paragraphTextOf(session.part(), session.paragraphIds()[0]!) ?? '';

const insert = (session: TreeDocxSession, text: string) => ({
  op: 'insertText' as const,
  paragraphId: session.paragraphIds()[0]!,
  offset: 3,
  text,
});

/** Refuses any body whose text holds `forbidden`, and records the parts it was asked about. */
function refusing(forbidden: string) {
  const asked: OoxmlPart[] = [];
  const admission: WriteAdmission = ({ part }) => {
    asked.push(part);
    const text = JSON.stringify(part.root);
    return text.includes(forbidden) ? `no ${forbidden} here` : null;
  };
  return { asked, admission };
}

describe('the write admission', () => {
  test('refuses a local write over the part as it would be published, in its own words', () => {
    const { session, admit } = opened();
    const host = refusing('X');
    admit(host.admission);
    const revision = session.packageRevision();

    const result = session.applyTreeOps([insert(session, 'X')], undefined, undefined, BODY);

    expect(result).toEqual({ committed: false, rejected: true, opCount: 1, reason: 'no X here' });
    expect(textOf(session)).toBe('abc');
    expect(session.packageRevision()).toBe(revision);
    expect(session.canUndo()).toBe(false);
    expect(host.asked.map((part) => part.name)).toEqual(['/word/document.xml']);
    expect(paragraphTextOf(host.asked[0]!, session.paragraphIds()[0]!)).toBe('abcX');
  });

  test('admits what it does not refuse, and nothing is asked without an admission', () => {
    const { session, admit } = opened();
    admit(refusing('X').admission);

    expect(session.applyTreeOps([insert(session, 'Y')], undefined, undefined, BODY).committed).toBe(
      true
    );
    admit(undefined);
    expect(session.applyTreeOps([insert(session, 'X')], undefined, undefined, BODY).committed).toBe(
      true
    );
    expect(textOf(session)).toBe('abcXY');
  });

  test('a remote write, an undo and a redo are not asked', () => {
    const { session, admit } = opened();
    session.applyTreeOps([insert(session, 'X')], undefined, undefined, BODY);
    const host = refusing('X');
    admit(host.admission);

    session.undo();
    expect(textOf(session)).toBe('abc');
    session.redo();
    expect(textOf(session)).toBe('abcX');
    const remote = session.applyTreeOps([insert(session, 'X')], undefined, undefined, BODY, {
      origin: ORIGIN_IDS.mutationRemote,
    });

    expect(remote.committed).toBe(true);
    expect(textOf(session)).toBe('abcXX');
    expect(host.asked).toEqual([]);
  });

  test('a note inserted through the lifecycle lane is asked too', () => {
    const { session, admit } = opened();
    const host = refusing('footnoteReference');
    admit(host.admission);
    const revision = session.packageRevision();

    const result = session.applyTreeOps(
      [
        {
          op: 'insertNote',
          noteKind: 'footnote',
          paragraphId: session.paragraphIds()[0]!,
          offset: 1,
        },
      ],
      undefined,
      undefined,
      BODY
    );

    expect(result.committed).toBe(false);
    expect(result.reason).toBe('no footnoteReference here');
    expect(session.packageRevision()).toBe(revision);
  });
});
