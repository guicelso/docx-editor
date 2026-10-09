// Content inserted at a field's edge takes a run of its own: the runs that carry `w:fldChar` and
// `w:instrText` stay the field's, and a consumer that replaces the field by its runs keeps the text.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, type OoxmlNode, type OoxmlPart } from '../package/ooxml-tree.ts';
import { serializeOoxmlPart } from '../package/ooxml-serialize.ts';
import type { OoxmlParagraphNode } from '../package/ooxml-tree.ts';
import { createNodeIdAllocator } from '../package/ooxml-edit.ts';
import { storyBlocks } from '../../layout/story-roots.ts';
import { applyTreeOp } from '../store/tree-op-apply.ts';
import { insertRunPayloadAtOffset } from '../store/tree-op-insert-offset.ts';
import { validateTreeOp } from '../store/tree-op-validate.ts';
import type { TreeDocOp } from '../store/tree-op-types.ts';
import { WML_NAMESPACE_URI } from '../package/ooxml-shared.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const BOLD = '<w:rPr><w:b/></w:rPr>';
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const sdt = (inner: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="span:1"/><w:richText/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
const field = (rPr = '') =>
  `<w:r>${rPr}<w:fldChar w:fldCharType="begin"/></w:r>` +
  `<w:r>${rPr}<w:instrText xml:space="preserve"> MERGEFIELD "x" </w:instrText></w:r>` +
  `<w:r>${rPr}<w:fldChar w:fldCharType="separate"/></w:r>` +
  `<w:r>${rPr}<w:t>«x»</w:t></w:r>` +
  `<w:r>${rPr}<w:fldChar w:fldCharType="end"/></w:r>`;

function documentOf(body: string): OoxmlPart {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body><w:p>${body}</w:p></w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'app/xml' }
  );
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

function paragraphOf(part: OoxmlPart): OoxmlParagraphNode {
  return storyBlocks(part)[0] as OoxmlParagraphNode;
}

function controlId(part: OoxmlPart): string {
  const control = paragraphOf(part).children.find((child) => child.kind === 'contentControl');
  if (!control) throw new Error('no control');
  return control.id;
}

/** Apply, and confirm that validation answers what the applier did. */
function applied(part: OoxmlPart, op: TreeDocOp): OoxmlPart {
  const rejection = validateTreeOp(part, op);
  const result = applyTreeOp(part, op);
  if (!result.ok) throw new Error(`refused: ${result.reason}`);
  expect(rejection).toBeNull();
  return result.part;
}

function typed(body: string, offset: number, inside?: (part: OoxmlPart) => string): string {
  const part = documentOf(body);
  return serializeOoxmlPart(
    applied(part, {
      op: 'insertText',
      paragraphId: paragraphOf(part).id,
      offset,
      text: ', ',
      ...(inside ? { inside: inside(part) } : {}),
    })
  );
}

/** Every run that carries a `w:fldChar` or a `w:instrText` carries nothing else but its properties. */
function chromeRunsArePure(xml: string): boolean {
  const runs = xml.match(/<w:r>.*?<\/w:r>/gu) ?? [];
  return runs
    .filter((each) => each.includes('<w:fldChar') || each.includes('<w:instrText'))
    .every((each) => !each.includes('<w:t') && !each.includes('<w:tab'));
}

describe('content inserted at a field edge takes a run of its own', () => {
  test('typed after a field that ends the paragraph', () => {
    const xml = typed(field(), 1);
    expect(chromeRunsArePure(xml)).toBe(true);
    expect(xml).toContain(
      '<w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve">, </w:t></w:r>'
    );
  });

  test('typed before a field that starts the paragraph', () => {
    const xml = typed(field(), 0);
    expect(chromeRunsArePure(xml)).toBe(true);
    expect(xml).toContain(
      '<w:r><w:t xml:space="preserve">, </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/>'
    );
  });

  test('typed between two fields', () => {
    const xml = typed(field() + field(), 1);
    expect(chromeRunsArePure(xml)).toBe(true);
    expect(xml).toContain(
      '<w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve">, </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/>'
    );
  });

  test('typed after a field that ends a named control, inside it', () => {
    const xml = typed(sdt(run('CPF ') + field()), 5, controlId);
    expect(chromeRunsArePure(xml)).toBe(true);
    expect(xml).toContain(
      '<w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve">, </w:t></w:r></w:sdtContent>'
    );
  });

  test('the run keeps the formatting the field showed', () => {
    const xml = typed(field(BOLD), 1);
    expect(xml).toContain(
      `${BOLD}<w:fldChar w:fldCharType="end"/></w:r><w:r>${BOLD}<w:t xml:space="preserve">, </w:t></w:r>`
    );
  });

  test('typed after ordinary text beside a field still joins that text', () => {
    const xml = typed(run('CPF') + field(), 3);
    expect(xml).toContain('<w:t xml:space="preserve">CPF, </w:t>');
  });

  test('a payload placed at a field edge takes a run of its own too', () => {
    const part = documentOf(field());
    const paragraph = paragraphOf(part);
    const nextId = createNodeIdAllocator(part);
    const tab = {
      id: nextId(),
      kind: 'tab',
      namespaceUri: WML_NAMESPACE_URI,
      localName: 'tab',
      prefix: 'w',
      namespaceBindings: [],
      attributes: [],
      children: [],
    } as unknown as OoxmlNode;
    const result = insertRunPayloadAtOffset(part, paragraph, 1, [tab]);
    if (!result.ok) throw new Error('refused');
    const xml = serializeOoxmlPart(result.part);
    expect(chromeRunsArePure(xml)).toBe(true);
    expect(xml).toContain('<w:fldChar w:fldCharType="end"/></w:r><w:r><w:tab/></w:r>');
  });
});
