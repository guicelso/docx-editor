// The inline content-control edges at one offset, in reading order.
//
// Where controls start or end at one offset, the edges there name each place a caret can stand:
// a slot is the gap between two of them, drawn as tags or not.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, type OoxmlNode, type OoxmlPart } from '../package/ooxml-tree.ts';
import { isContentControl } from '../package/content-control-walk.ts';
import { contentControlEdgesAt } from '../store/content-control-edges.ts';
import { storyBlocks } from '../../layout/story-roots.ts';
import { contentControlTagSubjectOf } from '../../layout/content-control-tags.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const sdt = (tag: string, inner: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w:richText/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

function documentOf(body: string): OoxmlPart {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body><w:p>${body}</w:p></w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'app/xml' }
  );
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

/** Each edge at `offset` as its control's tag and an arrow: `G▸` opens G, `G◂` closes it. */
function edges(body: string, offset: number): readonly string[] {
  const part = documentOf(body);
  const tags = new Map<string, string>();
  const walk = (nodes: readonly OoxmlNode[]): void => {
    for (const node of nodes) {
      if (node.kind === 'textValue') continue;
      if (isContentControl(node)) tags.set(node.id, contentControlTagSubjectOf(node).tag ?? '?');
      walk(node.children);
    }
  };
  walk([part.root]);
  const paragraph = storyBlocks(part)[0]!;
  if (paragraph.kind !== 'paragraph') throw new Error('not a paragraph');
  return contentControlEdgesAt(part, paragraph, offset).map(
    ({ controlId, edge }) => `${tags.get(controlId)}${edge === 'open' ? '▸' : '◂'}`
  );
}

const GROUP = `${run('CPF ')}${sdt('G', sdt('B', run('RG')) + sdt('E', run('CNH')))}`;

describe('the control edges at an offset', () => {
  test('in reading order: closings before the openings that follow them', () => {
    expect(edges(GROUP, 4)).toEqual(['G▸', 'B▸']);
    expect(edges(GROUP, 6)).toEqual(['B◂', 'E▸']);
    expect(edges(GROUP, 9)).toEqual(['E◂', 'G◂']);
    expect(edges(GROUP, 5)).toEqual([]);
  });

  test('empty controls at one offset nest in the order they are written', () => {
    expect(edges(`${run('CPF ')}${sdt('G', sdt('B', '') + sdt('E', ''))}`, 4)).toEqual([
      'G▸',
      'B▸',
      'B◂',
      'E▸',
      'E◂',
      'G◂',
    ]);
  });

  test('a control inside a hyperlink is found through it', () => {
    expect(
      edges(`${run('a')}<w:hyperlink w:anchor="x">${sdt('B', run('RG'))}</w:hyperlink>`, 1)
    ).toEqual(['B▸']);
  });
});
