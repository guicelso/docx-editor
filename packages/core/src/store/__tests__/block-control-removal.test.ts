// Removing a block-level control WITH its content is a block removal, and keeps the invariants
// `deleteBlock` keeps: a story keeps a paragraph, a cell ends with one, and a section mark is not
// dropped with the paragraph that carries it. Keeping the content removes no block.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, type OoxmlNode, type OoxmlPart } from '../package/ooxml-tree.ts';
import { contentControlPropertiesOf } from '../package/content-control-nodes.ts';
import { applyTreeOp } from '../store/tree-op-apply.ts';
import { validateTreeOp } from '../store/tree-op-validate.ts';
import type { TreeDocOp } from '../store/tree-op-types.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const paragraph = (text: string, pPr = '') =>
  `<w:p>${pPr}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const block = (inner: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="B"/><w:richText/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
const cell = (inner: string) =>
  `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc>${inner}</w:tc></w:tr></w:tbl>`;
const sectionMark = '<w:pPr><w:sectPr/></w:pPr>';

function documentOf(body: string): OoxmlPart {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'app/xml' }
  );
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

function controlOf(part: OoxmlPart): string {
  let found: string | null = null;
  const walk = (node: OoxmlNode): void => {
    if (found || node.kind === 'textValue') return;
    if (node.kind === 'contentControl' && contentControlPropertiesOf(node).tag === 'B') {
      found = node.id;
      return;
    }
    node.children.forEach(walk);
  };
  walk(part.root);
  if (!found) throw new Error('no control');
  return found;
}

function refusalOf(body: string, keepContent: boolean): string | null {
  const part = documentOf(body);
  const op: TreeDocOp = { op: 'removeContentControl', controlId: controlOf(part), keepContent };
  const rejection = validateTreeOp(part, op);
  const result = applyTreeOp(part, op);
  expect(result.ok ? null : result.reason).toBe(rejection);
  return rejection;
}

describe('removing a block control with its content', () => {
  test('is a removal like any other where a paragraph survives', () => {
    expect(refusalOf(paragraph('a') + block(paragraph('b')), false)).toBeNull();
  });

  test('keeps a paragraph in the story', () => {
    expect(refusalOf(block(paragraph('a') + paragraph('b')), false)).toBe('block-required');
    expect(refusalOf(block(paragraph('a') + paragraph('b')), true)).toBeNull();
  });

  test('keeps the paragraph a cell ends with', () => {
    expect(refusalOf(cell(block(paragraph('y'))) + paragraph('c'), false)).toBe('block-required');
    expect(
      refusalOf(cell(block(paragraph('y')) + paragraph('x')) + paragraph('c'), false)
    ).toBeNull();
  });

  test('does not drop a section mark with its paragraph', () => {
    expect(
      refusalOf(paragraph('a') + block(paragraph('b', sectionMark)) + paragraph('c'), false)
    ).toBe('carries-section-mark');
  });
});
