// A complex field is one place in the offset model, spelled over several runs. A split at either of
// its edges must leave every one of those runs on the field's side: a field cut between its
// `begin` and its instruction reads as nothing on both halves.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, validateOoxmlPart, type OoxmlPart } from '../package/ooxml-tree.ts';
import { serializeOoxmlPart } from '../package/ooxml-serialize.ts';
import { applyTreeOp, type TreeDocOp } from '../store/tree-ops.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const FIELD =
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> MERGEFIELD x </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>val</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

function load(paragraph: string): OoxmlPart {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body><w:p>${paragraph}</w:p><w:sectPr/></w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'app/xml' }
  );
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

function firstParagraphId(part: OoxmlPart): string {
  const body = part.root.kind === 'textValue' ? null : part.root.children[0];
  const paragraph = body && body.kind !== 'textValue' ? body.children[0] : undefined;
  if (!paragraph) throw new Error('no paragraph');
  return paragraph.id;
}

function applied(part: OoxmlPart, op: TreeDocOp): string {
  const result = applyTreeOp(part, op);
  if (!result.ok) throw new Error(result.reason);
  expect(validateOoxmlPart(result.part).ok).toBe(true);
  return serializeOoxmlPart(result.part);
}

/** Each paragraph's runs, with the field's pieces named, so a cut field shows. */
function paragraphsOf(xml: string): string[] {
  return [...xml.matchAll(/<w:p>(.*?)<\/w:p>/g)].map((match) =>
    match[1]!
      .replace(/<w:fldChar w:fldCharType="(\w+)"\/>/g, '[$1]')
      .replace(/<w:instrText[^>]*>[^<]*<\/w:instrText>/g, '[instr]')
      .replace(/<[^>]+>/g, '')
  );
}

const WHOLE_FIELD = '[begin][instr][separate]val[end]';

describe('a split at the edge of a complex field', () => {
  test('after the field, the field stays in the head', () => {
    const part = load(`${run('ab')}${FIELD}${run('cd')}`);

    const xml = applied(part, {
      op: 'splitParagraph',
      paragraphId: firstParagraphId(part),
      offset: 3,
    });

    expect(paragraphsOf(xml)).toEqual([`ab${WHOLE_FIELD}`, 'cd']);
  });

  test('before the field, the field goes whole to the tail', () => {
    const part = load(`${run('ab')}${FIELD}${run('cd')}`);

    const xml = applied(part, {
      op: 'splitParagraph',
      paragraphId: firstParagraphId(part),
      offset: 2,
    });

    expect(paragraphsOf(xml)).toEqual(['ab', `${WHOLE_FIELD}cd`]);
  });

  test('a many-way split keeps the field whole the same way', () => {
    const part = load(`${run('ab')}${FIELD}${run('cd')}`);

    const xml = applied(part, {
      op: 'splitParagraphMany',
      paragraphId: firstParagraphId(part),
      offsets: [1, 3],
    });

    expect(paragraphsOf(xml)).toEqual(['a', `b${WHOLE_FIELD}`, 'cd']);
  });

  test('a control inserted right after the field lands after all of it', () => {
    const part = load(`${run('ab')}${FIELD}${run('cd')}`);

    const xml = applied(part, {
      op: 'insertInlineContentControl',
      paragraphId: firstParagraphId(part),
      offset: 3,
      tag: 'n1',
      text: 'x',
    });

    expect(paragraphsOf(xml)).toEqual([`ab${WHOLE_FIELD}xcd`]);
    expect(xml.indexOf('<w:fldChar w:fldCharType="end"/>')).toBeLessThan(xml.indexOf('<w:sdt>'));
  });
});
