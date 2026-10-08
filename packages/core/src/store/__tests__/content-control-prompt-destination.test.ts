// Text typed at a named place replaces that place's prompt, and no other.
//
// Two empty controls side by side share an offset, and each shows its prompt. The first
// keystroke into one replaces its prompt; the prompt of the neighbour sharing the offset stays.
// Beside a control, no prompt is replaced at all.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, type OoxmlNode, type OoxmlPart } from '../package/ooxml-tree.ts';
import { applyTreeOp } from '../store/tree-op-apply.ts';
import { validateTreeOp } from '../store/tree-op-validate.ts';
import type { TreeDocOp } from '../store/tree-op-types.ts';
import { contentControlPropertiesOf } from '../package/content-control-nodes.ts';
import {
  contentControlContentChildren,
  isContentControl,
} from '../package/content-control-walk.ts';
import { storyBlocks } from '../../layout/story-roots.ts';
import { contentControlTagSubjectOf } from '../../layout/content-control-tags.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const sdt = (tag: string, inner: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w:richText/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
const prompt = (tag: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w:showingPlcHdr/><w:richText/></w:sdtPr><w:sdtContent>${run('pp')}</w:sdtContent></w:sdt>`;

function documentOf(body: string): OoxmlPart {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body><w:p>${body}</w:p></w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'app/xml' }
  );
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

function bracketed(part: OoxmlPart): string {
  const paragraph = storyBlocks(part)[0]!;
  const walk = (nodes: readonly OoxmlNode[]): string =>
    nodes
      .map((node) => {
        if (node.kind === 'textValue') return node.value;
        if (isContentControl(node)) {
          return `${contentControlTagSubjectOf(node).tag}{${walk(contentControlContentChildren(node))}}`;
        }
        return walk(node.children);
      })
      .join('');
  return walk(paragraph.kind === 'textValue' ? [] : paragraph.children);
}

function controlId(part: OoxmlPart, tag: string): string {
  let found: string | null = null;
  const walk = (nodes: readonly OoxmlNode[]): void => {
    for (const node of nodes) {
      if (node.kind === 'textValue') continue;
      if (isContentControl(node) && contentControlTagSubjectOf(node).tag === tag) found = node.id;
      walk(node.children);
    }
  };
  walk([part.root]);
  if (found === null) throw new Error(`no control ${tag}`);
  return found;
}

function showsPrompt(part: OoxmlPart, tag: string): boolean {
  const walk = (nodes: readonly OoxmlNode[]): OoxmlNode | null => {
    for (const node of nodes) {
      if (node.kind === 'textValue') continue;
      if (isContentControl(node) && contentControlTagSubjectOf(node).tag === tag) return node;
      const inner = walk(node.children);
      if (inner) return inner;
    }
    return null;
  };
  return contentControlPropertiesOf(walk([part.root])!).showingPlaceholder;
}

function typed(
  part: OoxmlPart,
  offset: number,
  text: string,
  place: (part: OoxmlPart) => Partial<Extract<TreeDocOp, { op: 'insertText' }>>
): OoxmlPart {
  const op: TreeDocOp = {
    op: 'insertText',
    paragraphId: storyBlocks(part)[0]!.id,
    offset,
    text,
    ...place(part),
  };
  expect(validateTreeOp(part, op)).toBeNull();
  const result = applyTreeOp(part, op);
  if (!result.ok) throw new Error(`refused: ${result.reason}`);
  return result.part;
}

/** `CPF G{B{pp}E{pp}}`: B spans 4–6 and E 6–8, both showing their prompt. */
const EMPTY_BRANCHES = `${run('CPF ')}${sdt('G', prompt('B') + prompt('E'))}`;

describe('typing into a named control replaces only its prompt', () => {
  test('the first key into B replaces B, and the second stays in B', () => {
    const first = typed(documentOf(EMPTY_BRANCHES), 4, 'd', (part) => ({
      inside: controlId(part, 'B'),
    }));
    expect(bracketed(first)).toBe('CPF G{B{d}E{pp}}');
    expect(showsPrompt(first, 'E')).toBe(true);
    const second = typed(first, 5, 'e', (part) => ({ inside: controlId(part, 'B') }));
    expect(bracketed(second)).toBe('CPF G{B{de}E{pp}}');
  });

  test('a key in the middle of B’s prompt replaces B, never the neighbour sharing its end', () => {
    const typedIn = typed(documentOf(EMPTY_BRANCHES), 5, 'd', (part) => ({
      inside: controlId(part, 'B'),
    }));
    expect(bracketed(typedIn)).toBe('CPF G{B{d}E{pp}}');
  });
});

describe('typing beside a control replaces no prompt', () => {
  test('in front of a group whose first branch shows its prompt', () => {
    const before = typed(documentOf(EMPTY_BRANCHES), 4, 'X', (part) => ({
      beside: { controlId: controlId(part, 'G'), side: 'before' },
    }));
    expect(bracketed(before)).toBe('CPF XG{B{pp}E{pp}}');
    expect(showsPrompt(before, 'B')).toBe(true);
  });

  test('between a branch and a neighbour showing its prompt', () => {
    const between = typed(
      documentOf(`${run('CPF ')}${sdt('G', sdt('B', run('RG')) + prompt('E'))}`),
      6,
      'X',
      (part) => ({
        beside: { controlId: controlId(part, 'B'), side: 'after' },
      })
    );
    expect(bracketed(between)).toBe('CPF G{B{RG}XE{pp}}');
    expect(showsPrompt(between, 'E')).toBe(true);
  });
});

describe('with no place named, the offset rule stands', () => {
  test('a prompt at the edge is the first keystroke’s, as in Word', () => {
    const atEdge = typed(
      documentOf(`${run('CPF ')}${sdt('G', sdt('B', run('RG')) + prompt('E'))}`),
      6,
      'X',
      () => ({})
    );
    expect(bracketed(atEdge)).toBe('CPF G{B{RG}E{X}}');
  });
});
