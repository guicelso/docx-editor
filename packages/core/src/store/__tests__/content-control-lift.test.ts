// An inline control is lifted to the block level without a character changing: its paragraph is cut
// at the control's edges, the control keeps its identity, and the paragraph left holding it becomes
// the control's one paragraph, with the content's own nodes.

import { describe, expect, test } from 'bun:test';
import { contentControlPropertiesOf } from '../package/content-control-nodes.ts';
import { readOoxmlPart, type OoxmlNode, type OoxmlPart } from '../package/ooxml-tree.ts';
import { applyTreeOp } from '../store/tree-op-apply.ts';
import { treeOpReach } from '../store/tree-op-content-controls.ts';
import type { TreeDocOp } from '../store/tree-op-types.ts';
import { validateTreeOp } from '../store/tree-op-validate.ts';

type Element = Exclude<OoxmlNode, { kind: 'textValue' }>;

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const inline = (tag: string, inner: string, extra = '') =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/>${extra}<w:richText/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
const JUSTIFIED = '<w:pPr><w:jc w:val="both"/></w:pPr>';

function documentOf(body: string): OoxmlPart {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'app/xml' }
  );
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

function bodyOf(part: OoxmlPart): Element {
  const body = part.root.kind === 'textValue' ? undefined : part.root.children[0];
  if (!body || body.kind === 'textValue') throw new Error('no body');
  return body;
}

/** The body as `p(text)` for a paragraph and `tag{...}` for a block control, controls inside as `tag[...]`. */
function reading(part: OoxmlPart): string {
  const inlineText = (nodes: readonly OoxmlNode[]): string =>
    nodes
      .map((node) => {
        if (node.kind === 'textValue') return node.value;
        if (node.kind === 'contentControl') return `${tagOf(node)}[${inlineText(contentOf(node))}]`;
        if (node.kind === 'contentControlProperties' || node.kind === 'paragraphProperties')
          return '';
        return inlineText(node.children);
      })
      .join('');
  const blocks = (nodes: readonly OoxmlNode[]): string =>
    nodes
      .flatMap((node) => {
        if (node.kind === 'paragraph') return [`p(${inlineText(node.children)})`];
        if (node.kind === 'contentControl') return [`${tagOf(node)}{${blocks(contentOf(node))}}`];
        return [];
      })
      .join(' ');
  return blocks(bodyOf(part).children);
}

function tagOf(control: Element): string {
  return contentControlPropertiesOf(control).tag ?? '';
}

function contentOf(control: Element): readonly OoxmlNode[] {
  const content = control.children.find((child) => child.kind === 'contentControlContent');
  return content && content.kind !== 'textValue' ? content.children : [];
}

function controlNamed(part: OoxmlPart, tag: string): Element {
  let found: Element | null = null;
  const walk = (node: OoxmlNode): void => {
    if (found || node.kind === 'textValue') return;
    if (node.kind === 'contentControl' && tagOf(node) === tag) found = node;
    else node.children.forEach(walk);
  };
  walk(part.root);
  if (found === null) throw new Error(`no control ${tag}`);
  return found;
}

const idOf = (part: OoxmlPart, tag: string) => controlNamed(part, tag).id;

const lift = (controlId: string): TreeDocOp => ({ op: 'liftContentControl', controlId });

function applied(part: OoxmlPart, op: TreeDocOp): OoxmlPart {
  expect(validateTreeOp(part, op)).toBeNull();
  const result = applyTreeOp(part, op);
  if (!result.ok) throw new Error(`refused: ${result.reason}`);
  return result.part;
}

function refused(part: OoxmlPart, op: TreeDocOp): string {
  const rejection = validateTreeOp(part, op);
  const result = applyTreeOp(part, op);
  expect(result.ok).toBe(false);
  expect(rejection).toBe(result.ok ? null : result.reason);
  return rejection ?? '';
}

const LIST = inline('list', inline('item:a', run('A')) + inline('item:b', run('B')));

describe('liftContentControl', () => {
  test('the paragraph is cut at the control, which becomes a block control holding its content', () => {
    const part = documentOf(`<w:p>${JUSTIFIED}${run('Declara: ')}${LIST}${run(' fim.')}</w:p>`);
    const paragraphId = bodyOf(part).children[0]!.id;
    const ids = ['list', 'item:a', 'item:b'].map((tag) => idOf(part, tag));

    const lifted = applied(part, lift(idOf(part, 'list')));

    expect(reading(lifted)).toBe('p(Declara: ) list{p(item:a[A]item:b[B])} p( fim.)');
    expect(['list', 'item:a', 'item:b'].map((tag) => idOf(lifted, tag))).toEqual(ids);
    expect(bodyOf(lifted).children[0]!.id).toBe(paragraphId);
  });

  test('each paragraph the cut leaves keeps the paragraph properties, with an identity of its own', () => {
    const part = documentOf(
      `<w:p w14:paraId="0000000A">${JUSTIFIED}${run('a ')}${LIST}${run(' b')}</w:p>`
    );

    const lifted = applied(part, lift(idOf(part, 'list')));
    const paragraphs = paragraphsIn(bodyOf(lifted));

    expect(paragraphs.map(justification)).toEqual(['both', 'both', 'both']);
    expect(new Set(paragraphs.map(paraIdOf)).size).toBe(3);
  });

  test('a control alone in its paragraph takes the paragraph with it, and no empty paragraph is left', () => {
    const part = documentOf(`<w:p>${LIST}</w:p>`);
    const paragraphId = bodyOf(part).children[0]!.id;

    const lifted = applied(part, lift(idOf(part, 'list')));

    expect(reading(lifted)).toBe('list{p(item:a[A]item:b[B])}');
    expect(paragraphsIn(bodyOf(lifted)).map((paragraph) => paragraph.id)).toEqual([paragraphId]);
  });

  test('text on one side only cuts that side', () => {
    const part = documentOf(`<w:p>${run('Declara: ')}${LIST}</w:p>`);

    expect(reading(applied(part, lift(idOf(part, 'list'))))).toBe(
      'p(Declara: ) list{p(item:a[A]item:b[B])}'
    );
  });

  test('the controls it held are lifted in turn inside it, one paragraph each', () => {
    const part = documentOf(`<w:p>${run('Declara: ')}${LIST}</w:p>`);
    const list = applied(part, lift(idOf(part, 'list')));
    const first = applied(list, lift(idOf(list, 'item:a')));
    const both = applied(first, lift(idOf(first, 'item:b')));

    expect(reading(both)).toBe('p(Declara: ) list{item:a{p(A)} item:b{p(B)}}');
  });

  test('the control keeps showing its prompt, now in its one paragraph', () => {
    const prompt = `<w:sdt><w:sdtPr><w:tag w:val="item:a"/><w:showingPlcHdr/><w:richText/></w:sdtPr><w:sdtContent>${run('digite')}</w:sdtContent></w:sdt>`;
    const part = documentOf(`<w:p>${prompt}</w:p>`);

    const lifted = applied(part, lift(idOf(part, 'item:a')));

    expect(reading(lifted)).toBe('item:a{p(digite)}');
    expect(contentControlPropertiesOf(controlNamed(lifted, 'item:a')).showingPlaceholder).toBe(
      true
    );
  });

  test('a control held by another inline control, or already between blocks, is not lifted', () => {
    const nested = documentOf(`<w:p>${LIST}</w:p>`);
    const block = documentOf(
      `<w:sdt><w:sdtPr><w:tag w:val="box"/><w:richText/></w:sdtPr><w:sdtContent><w:p>${run('x')}</w:p></w:sdtContent></w:sdt>`
    );

    expect(refused(nested, lift(idOf(nested, 'item:a')))).toBe('unsupported');
    expect(refused(block, lift(idOf(block, 'box')))).toBe('unsupported');
  });

  test('a bound control, a locked one and a tracked lift are refused', () => {
    const bound = documentOf(
      `<w:p>${inline('x', run('a'), '<w:dataBinding w:xpath="/a" w:storeItemID="{1}"/>')}</w:p>`
    );
    const locked = documentOf(`<w:p>${inline('x', run('a'), '<w:lock w:val="sdtLocked"/>')}</w:p>`);
    const plain = documentOf(`<w:p>${LIST}</w:p>`);

    expect(refused(bound, lift(idOf(bound, 'x')))).toBe('bound');
    expect(refused(locked, lift(idOf(locked, 'x')))).toBe('locked');
    expect(
      refused(plain, {
        ...lift(idOf(plain, 'list')),
        revision: { author: 'A', date: '2026-10-11T00:00:00Z' },
      } as TreeDocOp)
    ).toBe('invalidArgs');
  });

  test('an unknown id and a node that is not a control are refused', () => {
    const part = documentOf(`<w:p>${LIST}</w:p>`);

    expect(refused(part, lift('nope'))).toBe('unknown-content-control');
    expect(refused(part, lift(bodyOf(part).children[0]!.id))).toBe('not-a-content-control');
  });

  test('what holds the control and what it holds answer for the lift', () => {
    const part = documentOf(`<w:p>${LIST}</w:p>`);
    const controlId = idOf(part, 'list');

    expect(treeOpReach(lift(controlId))).toEqual({
      kind: 'nodes',
      targets: [{ nodeId: controlId, structural: true }],
    });
  });
});

function paragraphsIn(node: Element): readonly Element[] {
  return node.children.flatMap((child): readonly Element[] => {
    if (child.kind === 'textValue') return [];
    return child.kind === 'paragraph' ? [child] : paragraphsIn(child);
  });
}

function justification(paragraph: Element): string | null {
  const properties = paragraph.children.find((child) => child.kind === 'paragraphProperties');
  const jc =
    properties && properties.kind !== 'textValue'
      ? properties.children.find((child) => child.kind !== 'textValue' && child.localName === 'jc')
      : undefined;
  return jc && jc.kind !== 'textValue'
    ? (jc.attributes.find((attribute) => attribute.localName === 'val')?.value ?? null)
    : null;
}

function paraIdOf(paragraph: Element): string | undefined {
  return paragraph.attributes.find((attribute) => attribute.localName === 'paraId')?.value;
}
