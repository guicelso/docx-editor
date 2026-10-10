// Backspace and Delete beside the tags of a BLOCK-level control. With the tags drawn, a key beside a
// tag selects the control whole and the next one removes it, a key in a control showing its prompt
// removes it at once, and nothing outside the control is ever taken. With the tags hidden, the start
// of the control's first paragraph is a paragraph boundary the key does not join across.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import type { OoxmlNode } from '../../store/package/ooxml-tree.ts';
import { contentControlPropertiesOf } from '../../store/package/content-control-nodes.ts';
import type { PaginatedSurface } from '../paginated-surface.ts';
import { mount } from './paginated-surface-fixtures.ts';

const paragraph = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const block = (tag: string, inner: string, extra = '') =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/>${extra}<w:richText/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

/** `a | B[b | c] | d`. */
const BODY = paragraph('a') + block('B', paragraph('b') + paragraph('c')) + paragraph('d');

const mounted: PaginatedSurface[] = [];
afterEach(() => {
  for (const surface of mounted.splice(0)) surface.destroy();
});

function open(body: string, tags: boolean): PaginatedSurface {
  const { surface } = mount(body);
  mounted.push(surface);
  if (tags) {
    surface.setContentControlView({
      tags: { labelsOf: ({ tag }) => ({ open: { text: `${tag}▸` }, close: { text: `◂${tag}` } }) },
    });
  }
  return surface;
}

function textOf(node: OoxmlNode): string {
  if (node.kind === 'textValue') return node.value;
  return node.children.map(textOf).join('');
}

/** The body's blocks as `text` and `tag[…]`, `|` between siblings. */
function reading(surface: PaginatedSurface): string {
  const blocks = (nodes: readonly OoxmlNode[]): string =>
    nodes
      .flatMap((node): string[] => {
        if (node.kind === 'paragraph') return [textOf(node)];
        if (node.kind !== 'contentControl') return [];
        const content = node.children.find((child) => child.kind === 'contentControlContent');
        const flag = contentControlPropertiesOf(node).showingPlaceholder ? '*' : '';
        const inner = content && content.kind !== 'textValue' ? content.children : [];
        return [`${contentControlPropertiesOf(node).tag}${flag}[${blocks(inner)}]`];
      })
      .join('|');
  const body = surface.session.part().root.children[0]!;
  return body.kind === 'textValue' ? '' : blocks(body.children);
}

function idOf(surface: PaginatedSurface, match: (node: OoxmlNode) => boolean): string {
  let found: string | null = null;
  const walk = (node: OoxmlNode): void => {
    if (found || node.kind === 'textValue') return;
    if (match(node)) found = node.id;
    else node.children.forEach(walk);
  };
  walk(surface.session.part().root);
  if (found === null) throw new Error('no such node');
  return found;
}

const paragraphNamed = (surface: PaginatedSurface, text: string) =>
  idOf(surface, (node) => node.kind === 'paragraph' && textOf(node) === text);
const controlNamed = (surface: PaginatedSurface, tag: string) =>
  idOf(
    surface,
    (node) => node.kind === 'contentControl' && contentControlPropertiesOf(node).tag === tag
  );

/** The caret at a paragraph offset, in the slot beside the named tag. */
function caretBeside(
  surface: PaginatedSurface,
  at: { readonly text: string; readonly offset: number },
  tag: string,
  edge: 'open' | 'close',
  side: 'before' | 'after'
): void {
  const caret = { paragraphId: paragraphNamed(surface, at.text), offset: at.offset };
  surface.setSelection(
    { anchor: caret, head: caret },
    { controlId: controlNamed(surface, tag), edge, side }
  );
}

function caretAt(surface: PaginatedSurface, text: string, offset: number): void {
  const caret = { paragraphId: paragraphNamed(surface, text), offset };
  surface.setSelection({ anchor: caret, head: caret });
}

/** The selection as `text:offset` for each end. */
function selection(surface: PaginatedSurface): string {
  const { anchor, head } = surface.state().selection;
  const name = (paragraphId: string) =>
    textOf(
      (function find(node: OoxmlNode): OoxmlNode {
        if (node.id === paragraphId) return node;
        for (const child of node.kind === 'textValue' ? [] : node.children) {
          const found = find(child);
          if (found.id === paragraphId) return found;
        }
        return node;
      })(surface.session.part().root)
    );
  return `${name(anchor.paragraphId)}:${anchor.offset}-${name(head.paragraphId)}:${head.offset}`;
}

describe('with the tags drawn, a key beside a block tag', () => {
  test('Backspace just inside the opening tag selects the control, and the next removes it', () => {
    const surface = open(BODY, true);
    caretBeside(surface, { text: 'b', offset: 0 }, 'B', 'open', 'after');
    surface.deleteBackward();
    expect(reading(surface)).toBe('a|B[b|c]|d');
    expect(selection(surface)).toBe('b:0-c:1');
    surface.deleteBackward();
    expect(reading(surface)).toBe('a|d');
    expect(selection(surface)).toBe('a:1-a:1');
  });

  test('Delete just inside the closing tag selects the control', () => {
    const surface = open(BODY, true);
    caretBeside(surface, { text: 'c', offset: 1 }, 'B', 'close', 'before');
    surface.deleteForward();
    expect(reading(surface)).toBe('a|B[b|c]|d');
    expect(selection(surface)).toBe('b:0-c:1');
  });

  test('a key in a block control showing its prompt removes it at once', () => {
    const surface = open(
      paragraph('a') + block('P', paragraph('digite'), '<w:showingPlcHdr/>') + paragraph('d'),
      true
    );
    caretAt(surface, 'digite', 0);
    surface.deleteBackward();
    expect(reading(surface)).toBe('a|d');
  });

  test('Backspace between two paragraphs of the control joins them', () => {
    const surface = open(BODY, true);
    caretAt(surface, 'c', 0);
    surface.deleteBackward();
    expect(reading(surface)).toBe('a|B[bc]|d');
  });
});

describe('with the tags hidden', () => {
  test('Backspace at the start of the control joins nothing: the caret goes to the paragraph before', () => {
    const surface = open(BODY, false);
    caretAt(surface, 'b', 0);
    surface.deleteBackward();
    expect(reading(surface)).toBe('a|B[b|c]|d');
    expect(selection(surface)).toBe('a:1-a:1');
  });

  test('Backspace between two paragraphs of the control joins them', () => {
    const surface = open(BODY, false);
    caretAt(surface, 'c', 0);
    surface.deleteBackward();
    expect(reading(surface)).toBe('a|B[bc]|d');
  });
});
