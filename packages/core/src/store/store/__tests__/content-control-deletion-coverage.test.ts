// An untracked deletion and the inline content controls it reaches: a control the range holds
// and reaches past goes with it, the control the range lies in shows its prompt when emptied,
// and a prompt the range reaches into comes back whole.

import { describe, expect, test } from 'bun:test';
import { contentControlPropertiesOf } from '../../package/content-control-nodes.ts';
import {
  contentControlContentChildren,
  isContentControl,
} from '../../package/content-control-walk.ts';
import { readOoxmlPart, type OoxmlNode, type OoxmlPart } from '../../package/ooxml-tree.ts';
import { storyBlocks } from '../../../layout/story-roots.ts';
import { applyTreeOp } from '../tree-op-apply.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const sdt = (tag: string, inner: string, extra = '') =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/>${extra}<w:richText/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
const prompt = (tag: string, text: string) => sdt(tag, run(text), '<w:showingPlcHdr/>');
const PROMPT_OPTIONS = { placeholderPrompt: () => 'type here' };

function partOf(body: string): OoxmlPart {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body><w:p>${body}</w:p></w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'app/xml' }
  );
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

/** The paragraph with each control written `tag{…}`, and `tag*{…}` while it shows its prompt. */
function bracketed(part: OoxmlPart): string {
  const paragraph = part.root.children[0]!;
  const walk = (nodes: readonly OoxmlNode[]): string =>
    nodes
      .map((node) => {
        if (node.kind === 'textValue') return node.value;
        if (isContentControl(node)) {
          const properties = contentControlPropertiesOf(node);
          const flag = properties.showingPlaceholder ? '*' : '';
          return `${properties.tag}${flag}{${walk(contentControlContentChildren(node))}}`;
        }
        if (node.localName === 'bookmarkStart') return '#';
        return node.localName === 'sdtPr' || node.localName === 'rPr' ? '' : walk(node.children);
      })
      .join('');
  return walk(paragraph.kind === 'textValue' ? [] : paragraph.children);
}

function deleted(body: string, start: number, end: number): string {
  const part = partOf(body);
  const paragraph = storyBlocks(part)[0]!;
  const result = applyTreeOp(
    part,
    { op: 'deleteText', paragraphId: paragraph.id, start, end },
    PROMPT_OPTIONS
  );
  if (!result.ok) throw new Error(result.reason);
  return bracketed(result.part);
}

/** `AB O{xyz} CD`: the control spans offsets 2 to 5. */
const OPTIONAL = `${run('AB')}${sdt('O', run('xyz'))}${run('CD')}`;
/** `A G{B{rg}E{cnh}} Z`: the group spans 1 to 6, its branches 1 to 3 and 3 to 6. */
const GROUP = `${run('A')}${sdt('G', sdt('B', run('rg')) + sdt('E', run('cnh')))}${run('Z')}`;

describe('a deletion that holds a control and reaches past it', () => {
  test('takes the control with it', () => {
    expect(deleted(OPTIONAL, 1, 6)).toBe('AD');
  });

  test('takes it when the range only reaches past one edge', () => {
    expect(deleted(OPTIONAL, 1, 5)).toBe('ACD');
    expect(deleted(OPTIONAL, 2, 6)).toBe('ABD');
  });

  test('takes a group and every branch in it', () => {
    expect(deleted(GROUP, 0, 7)).toBe('');
  });

  test('takes the branches it covers and leaves the group showing its prompt', () => {
    expect(deleted(GROUP, 1, 6)).toBe('AG*{type here}Z');
  });

  test('takes a covered branch and keeps the one it only reaches into', () => {
    expect(deleted(GROUP, 2, 6)).toBe('AG{B{r}}Z');
  });

  test('keeps a bookmark the covered control held', () => {
    const marked = `${run('AB')}${sdt('O', `<w:bookmarkStart w:id="1" w:name="m"/>${run('xyz')}`)}${run('CD')}`;
    expect(deleted(marked, 1, 6)).toBe('A#D');
  });

  test('keeps a wrapper locked against deletion, emptied and showing its prompt', () => {
    const locked = `${run('AB')}${sdt('O', run('xyz'), '<w:lock w:val="sdtLocked"/>')}${run('CD')}`;
    expect(deleted(locked, 1, 6)).toBe('AO*{type here}D');
  });

  test('takes a covered control showing its prompt', () => {
    expect(deleted(`${run('AB')}${prompt('O', 'pp')}${run('CD')}`, 1, 5)).toBe('AD');
  });

  test('leaves an empty shell at its edge, and takes one strictly inside it', () => {
    const shell = `${run('AB')}${sdt('S', '')}${run('CD')}`;
    expect(deleted(shell, 2, 3)).toBe('ABS{}D');
    expect(deleted(shell, 1, 3)).toBe('AD');
  });
});

describe('a deletion that lies inside one control', () => {
  test('empties it, and the control shows its prompt again', () => {
    expect(deleted(OPTIONAL, 2, 5)).toBe('ABO*{type here}CD');
  });

  test('leaves the rest of its content', () => {
    expect(deleted(OPTIONAL, 3, 4)).toBe('ABO{xz}CD');
  });
});

describe('a deletion that reaches into a prompt', () => {
  test('leaves the prompt whole instead of a few of its letters', () => {
    expect(deleted(`${run('AB')}${prompt('O', 'name')}${run('CD')}`, 2, 3)).toBe('ABO*{name}CD');
  });

  test('of the whole prompt leaves the prompt, not an empty control marked as one', () => {
    expect(deleted(`${run('AB')}${prompt('O', 'name')}${run('CD')}`, 2, 6)).toBe('ABO*{name}CD');
  });
});
