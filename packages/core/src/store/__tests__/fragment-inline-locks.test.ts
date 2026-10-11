// A fragment spliced into a line answers the lock question where it lands, as typing there does.
//
// The splice cuts the host paragraph and joins it back. That join edits nothing the document
// held, so a locked control elsewhere in the paragraph, or one the fragment carries, must not
// refuse a landing that typing at the same offset is allowed.

import { describe, expect, test } from 'bun:test';
import {
  bodyStoryRoot,
  contentControlPropertiesOf,
  contentControlTextOf,
  contentControlsIn,
  readOoxmlPart,
  storyParagraphs,
  type OoxmlNode,
  type OoxmlPart,
} from '../index.ts';
import { applyTreeOp } from '../store/tree-op-apply.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function parse(xml: string, name = '/word/document.xml'): OoxmlPart {
  const result = readOoxmlPart(xml, { name, contentType: 'application/xml' });
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

const documentOf = (body: string): OoxmlPart =>
  parse(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`);

const fragmentOf = (inner: string): OoxmlNode =>
  parse(`<w:p xmlns:w="${W}">${inner}</w:p>`, '/fragment.xml').root;

const run = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const locked = (tag: string, held: string): string =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w:lock w:val="contentLocked"/></w:sdtPr>` +
  `<w:sdtContent>${run(held)}</w:sdtContent></w:sdt>`;

function paragraphs(part: OoxmlPart) {
  return storyParagraphs(bodyStoryRoot(part)!);
}

function textOf(node: OoxmlNode): string {
  return node.kind === 'textValue' ? node.value : node.children.map(textOf).join('');
}

/** Each control by tag: its own text and its lock. */
function controls(part: OoxmlPart) {
  return contentControlsIn(part.root).map(({ node }) => ({
    tag: contentControlPropertiesOf(node).tag,
    lock: contentControlPropertiesOf(node).lock,
    text: contentControlTextOf(node),
  }));
}

describe('insertFragment · inline landing and locks', () => {
  test('lands beside a locked control elsewhere in the paragraph, as typing there does', () => {
    const part = documentOf(`<w:p>${locked('cited', 'IGLOO')}${run(' e o vendedor')}</w:p>`);
    const paragraphId = paragraphs(part)[0]!.id;

    expect(applyTreeOp(part, { op: 'insertText', paragraphId, offset: 10, text: 'x' }).ok).toBe(
      true
    );
    const landed = applyTreeOp(part, {
      op: 'insertFragment',
      paragraphId,
      offset: 10,
      blocks: [fragmentOf(run('novo '))],
    });

    expect(landed.ok).toBe(true);
    if (!landed.ok) return;
    expect(paragraphs(landed.part)).toHaveLength(1);
    expect(textOf(paragraphs(landed.part)[0]!)).toBe('IGLOO e o novo vendedor');
    expect(controls(landed.part)).toEqual([{ tag: 'cited', lock: 'contentLocked', text: 'IGLOO' }]);
  });

  test('lands a fragment that carries a locked control, nested controls and locks intact', () => {
    const part = documentOf(`<w:p>${run('pelo vendedor')}</w:p>`);
    const paragraphId = paragraphs(part)[0]!.id;
    const carried =
      `<w:sdt><w:sdtPr><w:tag w:val="entry"/><w:lock w:val="contentLocked"/></w:sdtPr>` +
      `<w:sdtContent>${run('Pre: ')}${locked('cited', 'IGLOO')}</w:sdtContent></w:sdt>`;

    const landed = applyTreeOp(part, {
      op: 'insertFragment',
      paragraphId,
      offset: 5,
      blocks: [fragmentOf(carried)],
    });

    expect(landed.ok).toBe(true);
    if (!landed.ok) return;
    expect(textOf(paragraphs(landed.part)[0]!)).toBe('pelo Pre: IGLOOvendedor');
    expect(controls(landed.part)).toEqual([
      { tag: 'entry', lock: 'contentLocked', text: 'Pre: IGLOO' },
      { tag: 'cited', lock: 'contentLocked', text: 'IGLOO' },
    ]);
  });

  test('a landing inside locked content is still refused', () => {
    const part = documentOf(
      `<w:sdt><w:sdtPr><w:tag w:val="region"/><w:lock w:val="contentLocked"/></w:sdtPr>` +
        `<w:sdtContent><w:p>${run('held text')}</w:p></w:sdtContent></w:sdt>`
    );
    const paragraphId = paragraphs(part)[0]!.id;

    expect(
      applyTreeOp(part, {
        op: 'insertFragment',
        paragraphId,
        offset: 2,
        blocks: [fragmentOf(run('x'))],
      })
    ).toEqual({ ok: false, reason: 'locked' });
  });
});
