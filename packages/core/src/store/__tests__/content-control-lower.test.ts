// A block control holding one paragraph is lowered into that paragraph without a character changing:
// the paragraph stays, takes the control's place, and holds the control, which keeps its identity and
// holds what the paragraph held. It is the inverse of `liftContentControl`.

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
const lower = (controlId: string): TreeDocOp => ({ op: 'lowerContentControl', controlId });

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

const block = (tag: string, inner: string, extra = '') =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/>${extra}<w:richText/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

describe('lowerContentControl', () => {
  test("the one paragraph takes the control's place and holds it, both keeping their identity", () => {
    const part = documentOf(
      `<w:p>${run('Declara:')}</w:p>${block('list', `<w:p>${JUSTIFIED}${run('A')}</w:p>`)}<w:p>${run('fim.')}</w:p>`
    );
    const controlId = idOf(part, 'list');
    const paragraphId = contentOf(controlNamed(part, 'list'))[0]!.id;

    const lowered = applied(part, lower(controlId));

    expect(reading(lowered)).toBe('p(Declara:) p(list[A]) p(fim.)');
    expect(idOf(lowered, 'list')).toBe(controlId);
    expect(bodyOf(lowered).children[1]!.id).toBe(paragraphId);
    expect(justification(bodyOf(lowered).children[1] as Element)).toBe('both');
  });

  test('a list lifted and its items lifted come back down item by item, then the list, each paragraph joined', () => {
    const part = documentOf(
      `<w:p>${run('Declara: ')}${inline('list', inline('item:a', run('A')) + inline('item:b', run('B')))}</w:p>`
    );
    const lifted = [idOf(part, 'list'), 'item:a', 'item:b'].reduce(
      (current, target) =>
        applied(current, lift(target.startsWith('item') ? idOf(current, target) : target)),
      part
    );
    const items = applied(
      applied(lifted, lower(idOf(lifted, 'item:a'))),
      lower(idOf(lifted, 'item:b'))
    );
    const [first, second] = paragraphsIn(controlNamed(items, 'list'));
    const joined = applied(items, {
      op: 'joinParagraphs',
      firstId: first!.id,
      secondId: second!.id,
    });

    expect(reading(items)).toBe('p(Declara: ) list{p(item:a[A]) p(item:b[B])}');
    expect(reading(applied(joined, lower(idOf(joined, 'list'))))).toBe(
      'p(Declara: ) p(list[item:a[A]item:b[B]])'
    );
  });

  test('the control keeps showing its prompt, now in the line', () => {
    const prompt = `<w:sdt><w:sdtPr><w:tag w:val="item:a"/><w:showingPlcHdr/><w:richText/></w:sdtPr><w:sdtContent><w:p>${run('digite')}</w:p></w:sdtContent></w:sdt>`;
    const part = documentOf(prompt);

    const lowered = applied(part, lower(idOf(part, 'item:a')));

    expect(reading(lowered)).toBe('p(item:a[digite])');
    expect(contentControlPropertiesOf(controlNamed(lowered, 'item:a')).showingPlaceholder).toBe(
      true
    );
  });

  test('a control holding several paragraphs, or already in a line, is not lowered', () => {
    const several = documentOf(block('box', `<w:p>${run('a')}</w:p><w:p>${run('b')}</w:p>`));
    const inLine = documentOf(`<w:p>${inline('x', run('a'))}</w:p>`);

    expect(refused(several, lower(idOf(several, 'box')))).toBe('unsupported');
    expect(refused(inLine, lower(idOf(inLine, 'x')))).toBe('unsupported');
  });

  test('a bound control, a locked one and a tracked lowering are refused', () => {
    const one = `<w:p>${run('a')}</w:p>`;
    const bound = documentOf(block('x', one, '<w:dataBinding w:xpath="/a" w:storeItemID="{1}"/>'));
    const locked = documentOf(block('x', one, '<w:lock w:val="sdtLocked"/>'));
    const plain = documentOf(block('x', one));

    expect(refused(bound, lower(idOf(bound, 'x')))).toBe('bound');
    expect(refused(locked, lower(idOf(locked, 'x')))).toBe('locked');
    expect(
      refused(plain, {
        ...lower(idOf(plain, 'x')),
        revision: { author: 'A', date: '2026-10-11T00:00:00Z' },
      } as TreeDocOp)
    ).toBe('invalidArgs');
  });

  test('an unknown id and a node that is not a control are refused', () => {
    const part = documentOf(block('x', `<w:p>${run('a')}</w:p>`));

    expect(refused(part, lower('nope'))).toBe('unknown-content-control');
    expect(refused(part, lower(paragraphsIn(bodyOf(part))[0]!.id))).toBe('not-a-content-control');
  });

  test('what holds the control and what it holds answer for the lowering', () => {
    const part = documentOf(block('x', `<w:p>${run('a')}</w:p>`));
    const controlId = idOf(part, 'x');

    expect(treeOpReach(lower(controlId))).toEqual({
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
