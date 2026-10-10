// The edges of BLOCK-level content controls: a control opens at the start of its first paragraph
// and closes at the end of its last, outside every inline edge there, outer controls outside inner.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, type OoxmlNode, type OoxmlPart } from '../package/ooxml-tree.ts';
import { contentControlPropertiesOf } from '../package/content-control-nodes.ts';
import { blockControlEdgesOf } from '../store/block-control-edges.ts';
import { contentControlEdgesAt } from '../store/content-control-edges.ts';
import { applyTreeOp } from '../store/tree-op-apply.ts';
import type { OoxmlParagraphNode } from '../package/ooxml-tree.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const paragraph = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const inline = (tag: string, text: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/></w:sdtPr><w:sdtContent><w:r><w:t>${text}</w:t></w:r></w:sdtContent></w:sdt>`;
const block = (tag: string, inner: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w:richText/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

function documentOf(body: string): OoxmlPart {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'app/xml' }
  );
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

function textOf(node: OoxmlNode): string {
  if (node.kind === 'textValue') return node.value;
  return node.children.map(textOf).join('');
}

function paragraphNamed(part: OoxmlPart, text: string): OoxmlParagraphNode {
  let found: OoxmlParagraphNode | null = null;
  const walk = (node: OoxmlNode): void => {
    if (found || node.kind === 'textValue') return;
    if (node.kind === 'paragraph' && textOf(node) === text) {
      found = node as OoxmlParagraphNode;
      return;
    }
    node.children.forEach(walk);
  };
  walk(part.root);
  if (!found) throw new Error(`no paragraph ${text}`);
  return found;
}

const tagOf = (control: OoxmlNode) => contentControlPropertiesOf(control).tag;

/** The edges at an offset as `tag:open` / `tag:close`, in reading order. */
function edges(part: OoxmlPart, text: string, offset: number): string[] {
  const named = new Map<string, string>();
  const walk = (node: OoxmlNode): void => {
    if (node.kind === 'textValue') return;
    if (node.kind === 'contentControl') named.set(node.id, tagOf(node) ?? '?');
    node.children.forEach(walk);
  };
  walk(part.root);
  return contentControlEdgesAt(part, paragraphNamed(part, text), offset).map(
    (edge) => `${named.get(edge.controlId)}:${edge.edge}`
  );
}

describe('blockControlEdgesOf', () => {
  test('a control opens at its first paragraph and closes at its last', () => {
    const part = documentOf(paragraph('a') + block('B', paragraph('b') + paragraph('c')));
    const index = blockControlEdgesOf(part);

    expect(index.get(paragraphNamed(part, 'b').id)?.opens.map(tagOf)).toEqual(['B']);
    expect(index.get(paragraphNamed(part, 'b').id)?.closes).toEqual([]);
    expect(index.get(paragraphNamed(part, 'c').id)?.closes.map(tagOf)).toEqual(['B']);
    expect(index.get(paragraphNamed(part, 'a').id)).toBeUndefined();
  });

  test('a control that starts or ends with a table opens and closes at its edge cell paragraphs', () => {
    const cell = (inner: string) => `<w:tc>${inner}</w:tc>`;
    const table = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/><w:gridCol w:w="100"/></w:tblGrid><w:tr>${cell(paragraph('x1') + paragraph('x2'))}${cell(paragraph('y'))}</w:tr></w:tbl>`;
    const part = documentOf(
      paragraph('a') +
        block('T', table) +
        block(
          'M',
          paragraph('m') + table.replace('x1', 'z1').replace('x2', 'z2').replace('>y<', '>w<')
        )
    );
    const index = blockControlEdgesOf(part);

    expect(index.get(paragraphNamed(part, 'x1').id)?.opens.map(tagOf)).toEqual(['T']);
    expect(index.get(paragraphNamed(part, 'y').id)?.closes.map(tagOf)).toEqual(['T']);
    expect(index.get(paragraphNamed(part, 'm').id)?.opens.map(tagOf)).toEqual(['M']);
    expect(index.get(paragraphNamed(part, 'w').id)?.closes.map(tagOf)).toEqual(['M']);
  });

  test('nested controls open outer first and close inner first', () => {
    const part = documentOf(block('L', block('I1', paragraph('x')) + block('I2', paragraph('y'))));
    const index = blockControlEdgesOf(part);

    expect(index.get(paragraphNamed(part, 'x').id)?.opens.map(tagOf)).toEqual(['L', 'I1']);
    expect(index.get(paragraphNamed(part, 'x').id)?.closes.map(tagOf)).toEqual(['I1']);
    expect(index.get(paragraphNamed(part, 'y').id)?.opens.map(tagOf)).toEqual(['I2']);
    expect(index.get(paragraphNamed(part, 'y').id)?.closes.map(tagOf)).toEqual(['I2', 'L']);
  });

  test('an edit answers for the part it made, and the part before it still answers for itself', () => {
    const before = documentOf(paragraph('a') + paragraph('b'));
    const result = applyTreeOp(before, {
      op: 'wrapBlocksInContentControl',
      firstBlockId: paragraphNamed(before, 'a').id,
      lastBlockId: paragraphNamed(before, 'b').id,
      tag: 'W',
    });
    if (!result.ok) throw new Error(result.reason);

    expect(
      blockControlEdgesOf(result.part).get(paragraphNamed(result.part, 'a').id)?.opens.map(tagOf)
    ).toEqual(['W']);
    expect(blockControlEdgesOf(before).get(paragraphNamed(before, 'a').id)).toBeUndefined();
  });
});

describe('contentControlEdgesAt with block controls', () => {
  test('block edges stand outside the inline edges of the same paragraph', () => {
    const part = documentOf(block('B', `<w:p>${inline('i', 'x')}</w:p>`));

    expect(edges(part, 'x', 0)).toEqual(['B:open', 'i:open']);
    expect(edges(part, 'x', 1)).toEqual(['i:close', 'B:close']);
  });

  test('an empty paragraph holds every edge of a control that has nothing else', () => {
    const part = documentOf(paragraph('a') + block('B', '<w:p/>'));
    const empty = (() => {
      const controls = blockControlEdgesOf(part);
      for (const [paragraphId, at] of controls) if (at.opens.length > 0) return paragraphId;
      throw new Error('no block control');
    })();
    const named = contentControlEdgesAt(part, nodeById(part, empty), 0).map((edge) => edge.edge);

    expect(named).toEqual(['open', 'close']);
  });

  test('a paragraph in the middle of a control has no block edge', () => {
    const part = documentOf(block('B', paragraph('a') + paragraph('b') + paragraph('c')));

    expect(edges(part, 'b', 0)).toEqual([]);
    expect(edges(part, 'b', 1)).toEqual([]);
  });
});

function nodeById(part: OoxmlPart, id: string): OoxmlParagraphNode {
  let found: OoxmlParagraphNode | null = null;
  const walk = (node: OoxmlNode): void => {
    if (found || node.kind === 'textValue') return;
    if (node.id === id) {
      found = node as OoxmlParagraphNode;
      return;
    }
    node.children.forEach(walk);
  };
  walk(part.root);
  if (!found) throw new Error(`no node ${id}`);
  return found;
}
