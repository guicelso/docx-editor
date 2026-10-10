// The store side of a caret between tags: named, every slot's owner receives the typed text.
//
// A tagged edge shows several caret slots for ONE model offset. The surface already types into
// "the control the caret is in" by naming it (`inside`); this pins that the store can honour
// that name at every edge the tags make visible.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, type OoxmlNode, type OoxmlPart } from '../../store/package/ooxml-tree.ts';
import { applyTreeOp } from '../../store/store/tree-op-apply.ts';
import { storyBlocks } from '../story-roots.ts';
import { contentControlSubjectOf } from '../content-control-properties.ts';
import {
  isContentControl,
  contentControlContentChildren,
} from '../../store/package/content-control-walk.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const sdt = (tag: string, inner: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

function documentOf(): OoxmlPart {
  const body = `<w:p>${run('CPF ')}${sdt(
    'G',
    sdt('B', run('RG')) + sdt('E', run('CNH'))
  )}${run(' fim')}</w:p>`;
  const result = readOoxmlPart(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`, {
    name: '/word/document.xml',
    contentType: 'app/xml',
  });
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

/** The paragraph as text with its controls bracketed: `CPF G{B{RG}E{CNH}} fim`. */
function bracketed(part: OoxmlPart): string {
  const paragraph = storyBlocks(part)[0]!;
  const walk = (nodes: readonly OoxmlNode[]): string =>
    nodes
      .map((node) => {
        if (node.kind === 'textValue') return node.value;
        if (isContentControl(node)) {
          return `${contentControlSubjectOf(node).tag}{${walk(contentControlContentChildren(node))}}`;
        }
        return walk(node.children);
      })
      .join('');
  return walk(
    paragraph.kind === 'textValue'
      ? []
      : paragraph.children.filter((c) => c.kind !== 'paragraphProperties')
  );
}

function controlId(part: OoxmlPart, tag: string): string {
  let found: string | null = null;
  const walk = (nodes: readonly OoxmlNode[]): void => {
    for (const node of nodes) {
      if (node.kind === 'textValue') continue;
      if (isContentControl(node) && contentControlSubjectOf(node).tag === tag) found = node.id;
      walk(node.children);
    }
  };
  walk([part.root]);
  if (!found) throw new Error(`no control ${tag}`);
  return found;
}

function typed(
  offset: number,
  inside?: string,
  beside?: { readonly tag: string; readonly side: 'before' | 'after' }
): string {
  const part = documentOf();
  const paragraphId = storyBlocks(part)[0]!.id;
  const result = applyTreeOp(part, {
    op: 'insertText',
    paragraphId,
    offset,
    text: 'X',
    ...(inside === undefined ? {} : { inside: controlId(part, inside) }),
    ...(beside ? { beside: { controlId: controlId(part, beside.tag), side: beside.side } } : {}),
  });
  if (!result.ok) throw new Error(`refused: ${result.reason}`);
  return bracketed(result.part);
}

describe('typing at a tagged edge, slot by slot', () => {
  test('the document as it starts', () => {
    expect(bracketed(documentOf())).toBe('CPF G{B{RG}E{CNH}} fim');
  });

  test('offset 4 — with no request, the start edge belongs to the innermost control (today)', () => {
    expect(typed(4)).toBe('CPF G{B{XRG}E{CNH}} fim');
  });

  test('offset 4 — before [Grupo: beside G, before — outside, in front of the group', () => {
    expect(typed(4, undefined, { tag: 'G', side: 'before' })).toBe('CPF XG{B{RG}E{CNH}} fim');
  });

  test('offset 4 — between [Grupo and [Ramo 1: beside B, before — in the group, before the branch', () => {
    expect(typed(4, undefined, { tag: 'B', side: 'before' })).toBe('CPF G{XB{RG}E{CNH}} fim');
  });

  test('offset 6 — between ] and [Senão: beside B, after — in the group, between the branches', () => {
    expect(typed(6, undefined, { tag: 'B', side: 'after' })).toBe('CPF G{B{RG}XE{CNH}} fim');
  });

  test('offset 9 — after Grupo]: beside G, after', () => {
    expect(typed(9, undefined, { tag: 'G', side: 'after' })).toBe('CPF G{B{RG}E{CNH}}X fim');
  });

  test('a stale slot is refused, not landed elsewhere: beside G before at offset 6', () => {
    expect(() => typed(6, undefined, { tag: 'G', side: 'before' })).toThrow('refused');
  });

  test('offset 4 — after [Ramo 1: owner B, first character of the branch', () => {
    expect(typed(4, 'B')).toBe('CPF G{B{XRG}E{CNH}} fim');
  });

  test('offset 6 — before the branch close: owner B, last character of the branch', () => {
    expect(typed(6, 'B')).toBe('CPF G{B{RGX}E{CNH}} fim');
  });

  test('offset 6 — after [Senão: owner E, first character of the else', () => {
    expect(typed(6, 'E')).toBe('CPF G{B{RG}E{XCNH}} fim');
  });

  test('offset 9 — before the else close: owner E, last character of the else', () => {
    expect(typed(9, 'E')).toBe('CPF G{B{RG}E{CNHX}} fim');
  });

  test('offset 9 — after Grupo]: no owner, the store keeps it outside', () => {
    expect(typed(9)).toBe('CPF G{B{RG}E{CNH}}X fim');
  });
});

describe('a write beside a control answers to where it lands', () => {
  const locked = (tag: string, inner: string) =>
    `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w:lock w:val="sdtContentLocked"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
  const partOf = (body: string): OoxmlPart => {
    const result = readOoxmlPart(
      `<w:document xmlns:w="${W}"><w:body><w:p>${body}</w:p></w:body></w:document>`,
      { name: '/word/document.xml', contentType: 'app/xml' }
    );
    if (!result.ok) throw new Error(result.reason);
    return result.part;
  };
  const besideOf = (part: OoxmlPart, tag: string, side: 'before' | 'after', offset: number) =>
    applyTreeOp(part, {
      op: 'insertText',
      paragraphId: storyBlocks(part)[0]!.id,
      offset,
      text: 'X',
      beside: { controlId: controlId(part, tag), side },
    });

  test('beside a content-locked control is outside it, so the lock does not refuse', () => {
    const part = partOf(`${run('a ')}${locked('L', run('LOCK'))}`);
    const result = besideOf(part, 'L', 'before', 2);
    if (!result.ok) throw new Error(`refused: ${result.reason}`);
    expect(bracketed(result.part)).toBe('a XL{LOCK}');
  });

  test('beside a control held by a content-locked one is inside that one, and refused', () => {
    const part = partOf(`${run('a ')}${locked('P', sdt('C', run('IN')))}`);
    const result = besideOf(part, 'C', 'after', 4);
    expect(result.ok ? null : result.reason).toBe('locked');
  });

  test('naming an owner and a sibling at once is refused', () => {
    const part = partOf(`${run('a ')}${sdt('C', run('IN'))}`);
    const result = applyTreeOp(part, {
      op: 'insertText',
      paragraphId: storyBlocks(part)[0]!.id,
      offset: 2,
      text: 'X',
      inside: controlId(part, 'C'),
      beside: { controlId: controlId(part, 'C'), side: 'before' },
    });
    expect(result.ok ? null : result.reason).toBe('invalidArgs');
  });
});
