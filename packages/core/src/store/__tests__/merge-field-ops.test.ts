// A merge field a host authors: five runs at the caret's place, with the caret's formatting, and a
// rewrite that keeps the field's runs. The name is the host's; the instruction is the engine's.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, type OoxmlNode, type OoxmlPart } from '../package/ooxml-tree.ts';
import { serializeOoxmlPart } from '../package/ooxml-serialize.ts';
import { applyTreeOp } from '../store/tree-op-apply.ts';
import { validateTreeOp } from '../store/tree-op-validate.ts';
import type { TreeDocOp } from '../store/tree-op-types.ts';
import { mergeFieldNameOf, mergeFieldsOf } from '../store/tree-op-merge-fields.ts';
import {
  contentControlContentChildren,
  isContentControl,
} from '../package/content-control-walk.ts';
import { contentControlSubjectOf } from '../../layout/content-control-properties.ts';
import { storyBlocks } from '../../layout/story-roots.ts';
import type { OoxmlParagraphNode } from '../package/ooxml-tree.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const run = (text: string, rPr = '') => `<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const sdt = (tag: string, inner: string, properties = '') =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/>${properties}<w:richText/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
const field = (instruction: string, result: string, rPr = '') =>
  `<w:r>${rPr}<w:fldChar w:fldCharType="begin"/></w:r>` +
  `<w:r>${rPr}<w:instrText xml:space="preserve">${instruction}</w:instrText></w:r>` +
  `<w:r>${rPr}<w:fldChar w:fldCharType="separate"/></w:r>` +
  `<w:r>${rPr}<w:t>${result}</w:t></w:r>` +
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

/** The paragraph's text with each control bracketed by tag and each field as `«result»`. */
function reading(part: OoxmlPart): string {
  const walk = (nodes: readonly OoxmlNode[]): string =>
    nodes
      .map((node) => {
        if (node.kind === 'textValue') return node.value;
        if (node.kind === 'instrText') return '';
        if (isContentControl(node)) {
          return `${contentControlSubjectOf(node).tag ?? 'P'}{${walk(contentControlContentChildren(node))}}`;
        }
        if (node.kind === 'fldChar') {
          const type = node.attributes.find((attribute) => attribute.localName === 'fldCharType');
          return type?.value === 'begin' ? '«' : type?.value === 'end' ? '»' : '';
        }
        return walk(node.children);
      })
      .join('');
  return walk(paragraphOf(part).children);
}

function controlNamed(part: OoxmlPart, tag: string): OoxmlNode {
  let found: OoxmlNode | null = null;
  const walk = (nodes: readonly OoxmlNode[]): void => {
    for (const node of nodes) {
      if (node.kind === 'textValue') continue;
      if (isContentControl(node) && contentControlSubjectOf(node).tag === tag) found = node;
      walk(node.children);
    }
  };
  walk([part.root]);
  if (!found) throw new Error(`no control ${tag}`);
  return found;
}

/** Apply, and confirm that validation answers what the applier did. */
function applied(part: OoxmlPart, op: TreeDocOp): OoxmlPart {
  const rejection = validateTreeOp(part, op);
  const result = applyTreeOp(part, op, { placeholderPrompt: () => 'digite' });
  if (!result.ok) {
    expect(rejection).toBe(result.reason);
    throw new Error(`refused: ${result.reason}`);
  }
  expect(rejection).toBeNull();
  return result.part;
}

function refusalOf(part: OoxmlPart, op: TreeDocOp): string | null {
  const rejection = validateTreeOp(part, op);
  const result = applyTreeOp(part, op);
  expect(result.ok ? null : result.reason).toBe(rejection);
  return rejection;
}

const insert = (part: OoxmlPart, offset: number, extra: Partial<TreeDocOp> = {}): TreeDocOp =>
  ({
    op: 'insertMergeField',
    paragraphId: paragraphOf(part).id,
    offset,
    name: 'field:7c1f0a52-3b64-4d8e-9a10-2f5c6d7e8b90',
    result: 'X',
    ...extra,
  }) as TreeDocOp;

describe('insertMergeField', () => {
  test('writes five runs with the instruction spelled once, quoted', () => {
    const part = applied(documentOf(run('ab')), insert(documentOf(run('ab')), 1));
    const xml = serializeOoxmlPart(part);

    expect(reading(part)).toBe('a«X»b');
    expect(xml).toContain(
      '<w:instrText xml:space="preserve"> MERGEFIELD &quot;field:7c1f0a52-3b64-4d8e-9a10-2f5c6d7e8b90&quot; \\* MERGEFORMAT </w:instrText>'
    );
    expect(xml.match(/<w:r>/g)?.length).toBe(7);
  });

  test('every run carries the formatting of the text on the left', () => {
    const base = documentOf(run('Nome ', '<w:rPr><w:b/></w:rPr>') + run('fim'));
    const part = applied(base, insert(base, 5));
    const xml = serializeOoxmlPart(part);

    expect(reading(part)).toBe('Nome «X»fim');
    expect(xml.match(/<w:rPr><w:b\/><\/w:rPr>/g)?.length).toBe(6);
  });

  test('the atom occupies one offset, and the read reports its name and range', () => {
    const base = documentOf(run('ab'));
    const fields = mergeFieldsOf(paragraphOf(applied(base, insert(base, 1))));

    expect(fields).toEqual([
      expect.objectContaining({
        name: 'field:7c1f0a52-3b64-4d8e-9a10-2f5c6d7e8b90',
        start: 1,
        end: 2,
      }),
    ]);
  });

  test('beside a control it lands as the sibling the caller named', () => {
    const base = documentOf(run('a') + sdt('O', run('bc')) + run('d'));
    const control = controlNamed(base, 'O').id;

    expect(
      reading(applied(base, insert(base, 1, { beside: { controlId: control, side: 'before' } })))
    ).toBe('a«X»O{bc}d');
    expect(reading(applied(base, insert(base, 1, { inside: control })))).toBe('aO{«X»bc}d');
    expect(
      reading(applied(base, insert(base, 3, { beside: { controlId: control, side: 'after' } })))
    ).toBe('aO{bc}«X»d');
  });

  test('a caret in a control showing its prompt replaces the prompt', () => {
    const base = documentOf(run('a') + sdt('O', run('digite'), '<w:showingPlcHdr/>') + run('d'));

    expect(reading(applied(base, insert(base, 1, { inside: controlNamed(base, 'O').id })))).toBe(
      'aO{«X»}d'
    );
  });

  test('a locked control elsewhere in the paragraph does not refuse the field beside it', () => {
    const base = documentOf(
      run('a ') + sdt('L', run('fixo'), '<w:lock w:val="contentLocked"/>') + run(' b')
    );

    expect(reading(applied(base, insert(base, 1)))).toBe('a«X» L{fixo} b');
  });

  test('a locked control refuses the field written into it', () => {
    const base = documentOf(run('a') + sdt('L', run('fixo'), '<w:lock w:val="contentLocked"/>'));

    expect(refusalOf(base, insert(base, 2, { inside: controlNamed(base, 'L').id }))).toBe('locked');
  });

  test.each(['', 'with "quote"', 'back\\slash', 'line\nbreak'])('refuses the name %j', (name) => {
    const base = documentOf(run('ab'));

    expect(refusalOf(base, insert(base, 1, { name }))).toBe('invalidArgs');
  });
});

describe('setMergeField', () => {
  const NAME = 'field:7c1f0a52-3b64-4d8e-9a10-2f5c6d7e8b90';
  const bold = '<w:rPr><w:b/></w:rPr>';

  function rewritten(name: string, result: string): string {
    const base = documentOf(
      run('a') + field(` MERGEFIELD "${NAME}" \\* MERGEFORMAT `, '«nome»', bold) + run('b')
    );
    const [located] = mergeFieldsOf(paragraphOf(base));
    const part = applied(base, {
      op: 'setMergeField',
      paragraphId: paragraphOf(base).id,
      fieldNodeId: located!.fieldNodeId,
      name,
      result,
    } as TreeDocOp);
    return serializeOoxmlPart(part);
  }

  test('rewrites the name and the result, and the field keeps its runs and formatting', () => {
    const xml = rewritten('field:other', '«cpf»');

    expect(xml).toContain(' MERGEFIELD &quot;field:other&quot; \\* MERGEFORMAT ');
    expect(xml).toContain('«cpf»');
    expect(xml).not.toContain('«nome»');
    expect(xml.match(/<w:rPr><w:b\/><\/w:rPr>/g)?.length).toBe(5);
  });

  test('refuses a field that is not a merge field', () => {
    const base = documentOf(run('a') + field(' PAGE ', '1'));
    let begin: string | undefined;
    const walk = (node: OoxmlNode): void => {
      if (node.kind === 'textValue') return;
      if (begin === undefined && node.kind === 'fldChar') begin = node.id;
      node.children.forEach(walk);
    };
    walk(base.root);

    expect(
      refusalOf(base, {
        op: 'setMergeField',
        paragraphId: paragraphOf(base).id,
        fieldNodeId: begin!,
        name: NAME,
        result: 'x',
      } as TreeDocOp)
    ).toBe('invalidArgs');
  });
});

describe('mergeFieldNameOf', () => {
  test.each([
    [' MERGEFIELD "field:a@SELF|NUMERIC" \\* MERGEFORMAT ', 'field:a@SELF|NUMERIC'],
    [' MERGEFIELD Nome \\* MERGEFORMAT ', 'Nome'],
    [' PAGE ', null],
    [' MERGEFIELD "" ', null],
  ])('%j → %j', (instruction, name) => {
    expect(mergeFieldNameOf(instruction)).toBe(name);
  });
});
