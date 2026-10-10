// A block-level content control a host authors by identity: sibling blocks wrapped whole, or a new
// control at a place between blocks. Unwrapping is the removal every control already has.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPackage, writeOoxmlPackage } from '../package/ooxml-package.ts';
import { readOoxmlPart, type OoxmlNode, type OoxmlPart } from '../package/ooxml-tree.ts';
import { serializeOoxmlPart } from '../package/ooxml-serialize.ts';
import { contentControlPropertiesOf } from '../package/content-control-nodes.ts';
import { applyTreeOp } from '../store/tree-op-apply.ts';
import { validateTreeOp } from '../store/tree-op-validate.ts';
import type { TreeDocOp } from '../store/tree-op-types.ts';
import { openStore, transactBody, zipDoc } from './canonical-primitive-journal-coverage-support.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const TAG = 'span:optional:7c1f0a52-3b64-4d8e-9a10-2f5c6d7e8b90';

const paragraph = (id: string, text: string, pPr = '') =>
  `<w:p w14:paraId="${id}">${pPr}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const blockControl = (tag: string, inner: string, properties = '') =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/>${properties}<w:richText/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
const centered = '<w:pPr><w:jc w:val="center"/></w:pPr>';

function documentOf(body: string): OoxmlPart {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'app/xml' }
  );
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

/** Blocks as `text` for a paragraph and `tag[…]` for a block control, `|` between siblings. */
function reading(part: OoxmlPart): string {
  const body = part.root.kind === 'textValue' ? [] : part.root.children[0]!;
  const blocks = (nodes: readonly OoxmlNode[]): string =>
    nodes
      .flatMap((node): string[] => {
        if (node.kind === 'paragraph') return [textOf(node)];
        if (node.kind === 'contentControl') {
          const content = node.children.find((child) => child.kind === 'contentControlContent');
          const tag = contentControlPropertiesOf(node).tag ?? '?';
          return [
            `${tag}[${blocks(content?.kind === 'textValue' ? [] : (content?.children ?? []))}]`,
          ];
        }
        if (node.kind === 'table') return ['table'];
        return [];
      })
      .join('|');
  return body.kind === 'textValue' ? '' : blocks(body.children);
}

function textOf(node: OoxmlNode): string {
  if (node.kind === 'textValue') return node.value;
  return node.children.map(textOf).join('');
}

function nodeWhere(part: OoxmlPart, match: (node: OoxmlNode) => boolean): OoxmlNode {
  let found: OoxmlNode | null = null;
  const walk = (node: OoxmlNode): void => {
    if (found || node.kind === 'textValue') return;
    if (match(node)) {
      found = node;
      return;
    }
    node.children.forEach(walk);
  };
  walk(part.root);
  if (!found) throw new Error('no such node');
  return found;
}

const paragraphNamed = (part: OoxmlPart, text: string): OoxmlNode =>
  nodeWhere(part, (node) => node.kind === 'paragraph' && textOf(node) === text);
const controlNamed = (part: OoxmlPart, tag: string): OoxmlNode =>
  nodeWhere(
    part,
    (node) => node.kind === 'contentControl' && contentControlPropertiesOf(node).tag === tag
  );

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

const wrap = (part: OoxmlPart, first: string, last: string, extra = {}): TreeDocOp =>
  ({
    op: 'wrapBlocksInContentControl',
    firstBlockId: paragraphNamed(part, first).id,
    lastBlockId: paragraphNamed(part, last).id,
    tag: TAG,
    ...extra,
  }) as TreeDocOp;

const insertAt = (at: object, extra = {}): TreeDocOp =>
  ({ op: 'insertBlockContentControl', at, tag: TAG, ...extra }) as TreeDocOp;

describe('wrapBlocksInContentControl', () => {
  test('sibling paragraphs become the content of a new rich-text block control', () => {
    const base = documentOf(
      paragraph('1A', 'a') + paragraph('1B', 'b', centered) + paragraph('1C', 'c')
    );
    const part = applied(base, wrap(base, 'b', 'c', { lock: 'sdtLocked' }));
    const xml = serializeOoxmlPart(part);

    expect(reading(part)).toBe(`a|${TAG}[b|c]`);
    expect(xml).toContain(`<w:tag w:val="${TAG}"/>`);
    expect(xml).toContain('<w:lock w:val="sdtLocked"/>');
    expect(xml).toContain('<w:richText/>');
    expect(xml).toContain('<w:jc w:val="center"/>');
    expect(contentControlPropertiesOf(controlNamed(part, TAG)).showingPlaceholder).toBe(false);
  });

  test('the paragraphs keep their identity', () => {
    const base = documentOf(paragraph('1A', 'a') + paragraph('1B', 'b'));
    const before = paragraphNamed(base, 'b').id;

    expect(paragraphNamed(applied(base, wrap(base, 'a', 'b')), 'b').id).toBe(before);
  });

  test('a control named as both edges is wrapped itself; its children are wrapped inside it', () => {
    const base = documentOf(
      paragraph('1A', 'a') + blockControl('C', paragraph('1B', 'b') + paragraph('1C', 'c'))
    );
    const control = controlNamed(base, 'C').id;

    expect(
      reading(
        applied(base, {
          op: 'wrapBlocksInContentControl',
          firstBlockId: control,
          lastBlockId: control,
          tag: TAG,
        } as TreeDocOp)
      )
    ).toBe(`a|${TAG}[C[b|c]]`);
    expect(reading(applied(base, wrap(base, 'b', 'c')))).toBe(`a|C[${TAG}[b|c]]`);
  });

  test('blocks that are not siblings in order are refused', () => {
    const base = documentOf(
      paragraph('1A', 'a') + blockControl('C', paragraph('1B', 'b')) + paragraph('1C', 'c')
    );

    expect(refusalOf(base, wrap(base, 'a', 'b'))).toBe('not-adjacent-siblings');
    expect(refusalOf(base, wrap(base, 'c', 'a'))).toBe('not-adjacent-siblings');
  });

  test('a table is a block like any other: it is wrapped with its neighbours', () => {
    const table =
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>t</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const base = documentOf(paragraph('1A', 'a') + table + paragraph('1C', 'c'));
    const tableId = nodeWhere(base, (node) => node.kind === 'table').id;

    expect(reading(applied(base, wrap(base, 'a', 'c')))).toBe(`${TAG}[a|table|c]`);
    expect(
      reading(applied(base, wrap(base, 'a', 'a', { firstBlockId: tableId, lastBlockId: tableId })))
    ).toBe(`a|${TAG}[table]|c`);
  });

  test('an id that is no block is refused', () => {
    const base = documentOf(paragraph('1A', 'a'));
    const run = nodeWhere(base, (node) => node.kind === 'run').id;

    expect(refusalOf(base, wrap(base, 'a', 'a', { firstBlockId: 'missing' }))).toBe(
      'unknown-block'
    );
    expect(refusalOf(base, wrap(base, 'a', 'a', { lastBlockId: run }))).toBe('not-a-block');
  });

  test('a cell keeps the paragraph it must end with', () => {
    const table = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc>${paragraph('2A', 'x')}${paragraph('2B', 'y')}</w:tc></w:tr></w:tbl>`;
    const base = documentOf(table + paragraph('1C', 'c'));

    expect(refusalOf(base, wrap(base, 'x', 'y'))).toBe('block-required');
    expect(reading(applied(base, wrap(base, 'x', 'x')))).toBe('table|c');
  });

  test('a control whose content is locked refuses a wrap inside it; one elsewhere does not', () => {
    const locked = blockControl('L', paragraph('1B', 'b'), '<w:lock w:val="contentLocked"/>');
    const base = documentOf(paragraph('1A', 'a') + locked + paragraph('1C', 'c'));

    expect(refusalOf(base, wrap(base, 'b', 'b'))).toBe('locked');
    expect(reading(applied(base, wrap(base, 'c', 'c')))).toBe(`a|L[b]|${TAG}[c]`);
  });

  test('a tracked wrap is refused', () => {
    const base = documentOf(paragraph('1A', 'a'));

    expect(
      refusalOf(
        base,
        wrap(base, 'a', 'a', { revision: { author: 'QA', date: '2026-01-01T00:00:00Z' } })
      )
    ).toBe('invalidArgs');
  });

  test('metadata a control cannot carry is refused', () => {
    const base = documentOf(paragraph('1A', 'a'));

    expect(refusalOf(base, wrap(base, 'a', 'a', { tag: 'x'.repeat(4_097) }))).toBe(
      'invalid-property-value'
    );
    expect(refusalOf(base, wrap(base, 'a', 'a', { lock: 'everything' }))).toBe('invalidArgs');
  });
});

describe('insertBlockContentControl', () => {
  test('after a paragraph, a control showing its prompt in one paragraph', () => {
    const base = documentOf(paragraph('1A', 'a') + paragraph('1B', 'b'));
    const part = applied(base, insertAt({ after: paragraphNamed(base, 'a').id }));
    const control = controlNamed(part, TAG);

    expect(reading(part)).toBe(`a|${TAG}[digite]|b`);
    expect(contentControlPropertiesOf(control).showingPlaceholder).toBe(true);
  });

  test('before a block, holding the paragraphs given with fresh identities', () => {
    const base = documentOf(paragraph('1A', 'a') + paragraph('1B', 'b'));
    const given = documentOf(paragraph('1A', 'um', centered) + paragraph('1B', 'dois'));
    const blocks = [paragraphNamed(given, 'um'), paragraphNamed(given, 'dois')];
    const part = applied(base, insertAt({ before: paragraphNamed(base, 'b').id }, { blocks }));
    const xml = serializeOoxmlPart(part);

    expect(reading(part)).toBe(`a|${TAG}[um|dois]|b`);
    expect(contentControlPropertiesOf(controlNamed(part, TAG)).showingPlaceholder).toBe(false);
    expect(paragraphNamed(part, 'um').id).not.toBe(blocks[0]!.id);
    expect(xml).toContain('<w:jc w:val="center"/>');
    expect(xml.match(/w14:paraId="1A"/g)?.length).toBe(1);
    expect(xml.match(/w14:paraId="1B"/g)?.length).toBe(1);
  });

  test('inside a control showing its prompt, the new control replaces the prompt', () => {
    const owner = blockControl('O', paragraph('1B', 'digite'), '<w:showingPlcHdr/>');
    const base = documentOf(paragraph('1A', 'a') + owner);
    const part = applied(base, insertAt({ inside: controlNamed(base, 'O').id }));

    expect(reading(part)).toBe(`a|O[${TAG}[digite]]`);
    expect(contentControlPropertiesOf(controlNamed(part, 'O')).showingPlaceholder).toBe(false);
    expect(contentControlPropertiesOf(controlNamed(part, TAG)).showingPlaceholder).toBe(true);
  });

  test('inside a control that holds content is refused: its content has places of its own', () => {
    const base = documentOf(blockControl('O', paragraph('1B', 'b')));

    expect(refusalOf(base, insertAt({ inside: controlNamed(base, 'O').id }))).toBe('invalidArgs');
  });

  test('a place that names no block is refused', () => {
    const base = documentOf(paragraph('1A', 'a'));
    const run = nodeWhere(base, (node) => node.kind === 'run').id;

    expect(refusalOf(base, insertAt({ after: 'missing' }))).toBe('unknown-block');
    expect(refusalOf(base, insertAt({ before: run }))).toBe('not-a-block');
    expect(refusalOf(base, insertAt({ inside: run }))).toBe('not-a-content-control');
    expect(refusalOf(base, insertAt({ beside: paragraphNamed(base, 'a').id }))).toBe('invalidArgs');
  });

  test('after the last paragraph of a cell is refused: the cell must end with one', () => {
    const table = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc>${paragraph('2A', 'x')}</w:tc></w:tr></w:tbl>`;
    const base = documentOf(table + paragraph('1C', 'c'));

    expect(refusalOf(base, insertAt({ after: paragraphNamed(base, 'x').id }))).toBe(
      'block-required'
    );
  });

  test('the blocks given may be tables and block controls, each with fresh identities', () => {
    const base = documentOf(paragraph('1A', 'a'));
    const table = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc>${paragraph('2A', 'x')}</w:tc></w:tr></w:tbl>`;
    const given = documentOf(
      paragraph('1A', 'um') + table + blockControl('I', paragraph('3A', 'y'))
    );
    const body = given.root.kind === 'textValue' ? [] : given.root.children[0]!;
    const blocks =
      body.kind === 'textValue'
        ? []
        : body.children.filter(
            (node) =>
              node.kind === 'paragraph' || node.kind === 'table' || node.kind === 'contentControl'
          );
    const part = applied(base, insertAt({ after: paragraphNamed(base, 'a').id }, { blocks }));
    const xml = serializeOoxmlPart(part);

    expect(reading(part)).toBe(`a|${TAG}[um|table|I[y]]`);
    expect(xml.match(/w14:paraId="2A"/g)?.length ?? 0).toBeLessThanOrEqual(1);
    expect(xml.match(/w14:paraId="1A"/g)?.length).toBe(1);
  });

  test('the blocks given must be blocks, and at least one', () => {
    const base = documentOf(paragraph('1A', 'a'));
    const run = nodeWhere(base, (node) => node.kind === 'run');
    const inline = nodeWhere(
      documentOf(
        `<w:p><w:sdt><w:sdtPr><w:tag w:val="i"/></w:sdtPr><w:sdtContent><w:r><w:t>i</w:t></w:r></w:sdtContent></w:sdt></w:p>`
      ),
      (node) => node.kind === 'contentControl'
    );
    const at = { after: paragraphNamed(base, 'a').id };

    expect(refusalOf(base, insertAt(at, { blocks: [] }))).toBe('invalidArgs');
    expect(refusalOf(base, insertAt(at, { blocks: [run] }))).toBe('fragment-invalid-block');
    expect(refusalOf(base, insertAt(at, { blocks: [inline] }))).toBe('fragment-invalid-block');
  });

  test('a control whose content is locked refuses a place inside it', () => {
    const locked = blockControl('L', paragraph('1B', 'b'), '<w:lock w:val="contentLocked"/>');
    const base = documentOf(locked);

    expect(refusalOf(base, insertAt({ after: paragraphNamed(base, 'b').id }))).toBe('locked');
  });

  test('a tracked insertion is refused', () => {
    const base = documentOf(paragraph('1A', 'a'));

    expect(
      refusalOf(
        base,
        insertAt(
          { after: paragraphNamed(base, 'a').id },
          { revision: { author: 'QA', date: '2026-01-01T00:00:00Z' } }
        )
      )
    ).toBe('invalidArgs');
  });
});

describe('removeContentControl unwraps a block control', () => {
  const base = documentOf(
    paragraph('1A', 'a') + blockControl('C', paragraph('1B', 'b') + paragraph('1C', 'c'))
  );
  const remove = (keepContent: boolean): TreeDocOp => ({
    op: 'removeContentControl',
    controlId: controlNamed(base, 'C').id,
    keepContent,
  });

  test('keeping the content puts the paragraphs in its place, with their identity', () => {
    const part = applied(base, remove(true));

    expect(reading(part)).toBe('a|b|c');
    expect(paragraphNamed(part, 'b').id).toBe(paragraphNamed(base, 'b').id);
  });

  test('taking the content removes the paragraphs with it', () => {
    expect(reading(applied(base, remove(false)))).toBe('a');
  });
});

describe('a block control through the store', () => {
  const body = `${paragraph('1A', 'a')}${paragraph('1B', 'b', centered)}${paragraph('1C', 'c')}<w:sectPr/>`;

  test('a wrap is one undo step', () => {
    const store = openStore(zipDoc({ body }));
    const part = () => store.bodyStore().part;
    const before = reading(part());

    expect(transactBody(store, wrap(part(), 'b', 'c', { lock: 'sdtContentLocked' }))).toEqual({
      ok: true,
    });
    expect(reading(part())).toBe(`a|${TAG}[b|c]`);
    store.undo();
    expect(reading(part())).toBe(before);
  });

  test('saved and reopened, the tag, the lock, the paragraphs and their properties are intact', () => {
    const store = openStore(zipDoc({ body }));
    transactBody(store, wrap(store.bodyStore().part, 'b', 'c', { lock: 'sdtContentLocked' }));
    const reopened = readOoxmlPackage(writeOoxmlPackage(store.currentPackage()));
    if (!reopened.ok) throw new Error(reopened.reason);
    const part = reopened.package.parts.get(reopened.package.mainDocumentPart)!;
    const control = controlNamed(part, TAG);

    expect(reading(part)).toBe(`a|${TAG}[b|c]`);
    expect(contentControlPropertiesOf(control).lock).toBe('sdtContentLocked');
    expect(serializeOoxmlPart(part)).toContain('<w:jc w:val="center"/>');
  });
});
