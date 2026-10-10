// A paragraph written at a place between blocks: before or after a paragraph, a table or a
// block-level control. It takes the paragraph properties of the edge it stands beside, as a
// split takes the ones of the paragraph it divides, and it is where the caret goes.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, type OoxmlNode, type OoxmlPart } from '../package/ooxml-tree.ts';
import { serializeOoxmlPart } from '../package/ooxml-serialize.ts';
import { contentControlPropertiesOf } from '../package/content-control-nodes.ts';
import { applyTreeOp } from '../store/tree-op-apply.ts';
import { validateTreeOp } from '../store/tree-op-validate.ts';
import type { TreeDocOp } from '../store/tree-op-types.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const paragraph = (id: string, text: string, pPr = '', rPr = '') =>
  `<w:p w14:paraId="${id}">${pPr}<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const block = (tag: string, inner: string, extra = '') =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/>${extra}<w:richText/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
const indented = '<w:pPr><w:ind w:left="720"/></w:pPr>';
const bold = '<w:rPr><w:b/></w:rPr>';

function documentOf(body: string): OoxmlPart {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'app/xml' }
  );
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

function textOf(node: OoxmlNode): string {
  if (node.kind === 'textValue') return node.value;
  return node.children.map(textOf).join('');
}

function reading(part: OoxmlPart): string {
  const body = part.root.kind === 'textValue' ? null : part.root.children[0]!;
  const blocks = (nodes: readonly OoxmlNode[]): string =>
    nodes
      .flatMap((node): string[] => {
        if (node.kind === 'paragraph') return [textOf(node)];
        if (node.kind !== 'contentControl') return [];
        const content = node.children.find((child) => child.kind === 'contentControlContent');
        const inner = content && content.kind !== 'textValue' ? content.children : [];
        return [`${contentControlPropertiesOf(node).tag}[${blocks(inner)}]`];
      })
      .join('|');
  return body && body.kind !== 'textValue' ? blocks(body.children) : '';
}

function idOf(part: OoxmlPart, match: (node: OoxmlNode) => boolean): string {
  let found: string | null = null;
  const walk = (node: OoxmlNode): void => {
    if (found || node.kind === 'textValue') return;
    if (match(node)) found = node.id;
    else node.children.forEach(walk);
  };
  walk(part.root);
  if (found === null) throw new Error('no such node');
  return found;
}

const control = (part: OoxmlPart, tag: string) =>
  idOf(
    part,
    (node) => node.kind === 'contentControl' && contentControlPropertiesOf(node).tag === tag
  );
const paragraphNamed = (part: OoxmlPart, text: string) =>
  idOf(part, (node) => node.kind === 'paragraph' && textOf(node) === text);

function applied(part: OoxmlPart, op: TreeDocOp) {
  const rejection = validateTreeOp(part, op);
  const result = applyTreeOp(part, op);
  if (!result.ok) {
    expect(rejection).toBe(result.reason);
    throw new Error(`refused: ${result.reason}`);
  }
  expect(rejection).toBeNull();
  return result;
}

function refusalOf(part: OoxmlPart, op: TreeDocOp): string | null {
  const rejection = validateTreeOp(part, op);
  const result = applyTreeOp(part, op);
  expect(result.ok ? null : result.reason).toBe(rejection);
  return rejection;
}

const BODY =
  paragraph('1A', 'a') +
  block('B', paragraph('1B', 'b', indented, bold) + paragraph('1C', 'c', indented, bold)) +
  paragraph('1D', 'd');

describe('insertParagraph', () => {
  test('after a block control, a paragraph with the text, the last paragraph’s properties and face', () => {
    const base = documentOf(BODY);
    const result = applied(base, {
      op: 'insertParagraph',
      at: { after: control(base, 'B') },
      text: 'x',
    });
    const xml = serializeOoxmlPart(result.part);
    const created = paragraphNamed(result.part, 'x');

    expect(reading(result.part)).toBe('a|B[b|c]|x|d');
    expect(xml.match(/<w:ind w:left="720"\/>/g)?.length).toBe(3);
    expect(xml.match(/<w:b\/>/g)?.length).toBe(3);
    expect(result.ok && result.effect.caret?.paragraphId).toBe(created);
    expect(result.ok && result.effect.created).toEqual([created]);
  });

  test('before a block control, an empty paragraph with the first paragraph’s properties', () => {
    const base = documentOf(BODY);
    const result = applied(base, { op: 'insertParagraph', at: { before: control(base, 'B') } });

    expect(reading(result.part)).toBe('a||B[b|c]|d');
    expect(serializeOoxmlPart(result.part).match(/<w:ind w:left="720"\/>/g)?.length).toBe(3);
  });

  test('the new paragraph has an identity of its own', () => {
    const base = documentOf(BODY);
    const result = applied(base, {
      op: 'insertParagraph',
      at: { after: control(base, 'B') },
      text: 'x',
    });
    const ids = serializeOoxmlPart(result.part).match(/w14:paraId="[0-9A-F]+"/g) ?? [];

    expect(new Set(ids).size).toBe(ids.length);
    expect(ids.length).toBe(5);
  });

  test('a section mark stays on the paragraph that ends the section', () => {
    const base = documentOf(
      paragraph('1A', 'a', '<w:pPr><w:sectPr/></w:pPr>') + paragraph('1D', 'd')
    );
    const result = applied(base, {
      op: 'insertParagraph',
      at: { after: paragraphNamed(base, 'a') },
    });

    expect(serializeOoxmlPart(result.part).match(/<w:pPr><w:sectPr\/><\/w:pPr>/g)?.length).toBe(1);
  });

  test('a place inside a control or of no known shape is refused', () => {
    const base = documentOf(BODY);

    expect(refusalOf(base, { op: 'insertParagraph', at: { inside: control(base, 'B') } })).toBe(
      'invalidArgs'
    );
    expect(refusalOf(base, { op: 'insertParagraph', at: { after: 'missing' } })).toBe(
      'unknown-block'
    );
    expect(
      refusalOf(base, { op: 'insertParagraph', at: { after: control(base, 'B') }, text: '\u0001' })
    ).toBe('invalid-text');
    expect(
      refusalOf(base, { op: 'insertParagraph', at: { after: control(base, 'B') }, text: '' })
    ).toBe('invalid-text');
  });

  test('a control whose content is locked refuses a paragraph inside it', () => {
    const base = documentOf(block('L', paragraph('1B', 'b'), '<w:lock w:val="contentLocked"/>'));

    expect(
      refusalOf(base, { op: 'insertParagraph', at: { after: paragraphNamed(base, 'b') } })
    ).toBe('locked');
  });

  test('a tracked paragraph is refused', () => {
    const base = documentOf(BODY);

    expect(
      refusalOf(base, {
        op: 'insertParagraph',
        at: { after: control(base, 'B') },
        revision: { author: 'QA', date: '2026-01-01T00:00:00Z' },
      })
    ).toBe('invalidArgs');
  });
});
