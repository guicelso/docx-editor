// The slots outside a BLOCK control's tags are places between blocks. Typing there opens a
// paragraph with the text, Enter opens an empty one, and the caret goes into it; inside the
// control, Enter divides the paragraph within it. Pasting there is refused rather than landed
// inside the control.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import type { OoxmlNode } from '../../store/package/ooxml-tree.ts';
import { contentControlPropertiesOf } from '../../store/package/content-control-nodes.ts';
import type { PaginatedSurface } from '../paginated-surface.ts';
import { mount } from './paginated-surface-fixtures.ts';
import { slotPlacementOf } from '../surface-caret-slots.ts';

const paragraph = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const block = (tag: string, inner: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w:richText/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

/** `a | B[b | c] | d`. */
const BODY = paragraph('a') + block('B', paragraph('b') + paragraph('c')) + paragraph('d');

const mounted: PaginatedSurface[] = [];
afterEach(() => {
  for (const surface of mounted.splice(0)) surface.destroy();
});

function open(tags = true): PaginatedSurface {
  const { surface } = mount(BODY);
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

function reading(surface: PaginatedSurface): string {
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
const controlB = (surface: PaginatedSurface) =>
  idOf(
    surface,
    (node) => node.kind === 'contentControl' && contentControlPropertiesOf(node).tag === 'B'
  );

/** The caret at a paragraph offset, in the slot beside B's tag at that edge. */
function caretBeside(
  surface: PaginatedSurface,
  at: { readonly text: string; readonly offset: number },
  edge: 'open' | 'close',
  side: 'before' | 'after'
): void {
  const caret = { paragraphId: paragraphNamed(surface, at.text), offset: at.offset };
  surface.setSelection(
    { anchor: caret, head: caret },
    { controlId: controlB(surface), edge, side }
  );
}

const caretText = (surface: PaginatedSurface): string => {
  const { head } = surface.state().selection;
  return `${textOf(nodeById(surface, head.paragraphId))}:${head.offset}`;
};

function nodeById(surface: PaginatedSurface, id: string): OoxmlNode {
  let found: OoxmlNode | null = null;
  const walk = (node: OoxmlNode): void => {
    if (found || node.kind === 'textValue') return;
    if (node.id === id) found = node;
    else node.children.forEach(walk);
  };
  walk(surface.session.part().root);
  if (!found) throw new Error(`no node ${id}`);
  return found;
}

describe('the slot outside a block tag', () => {
  test('names the place between blocks: before the opening tag, after the closing one', () => {
    const block_ = { controlId: 'B', edge: 'open', level: 'block' } as const;

    expect(slotPlacementOf({ left: null, right: block_ })).toEqual({ block: { before: 'B' } });
    expect(slotPlacementOf({ left: { ...block_, edge: 'close' }, right: null })).toEqual({
      block: { after: 'B' },
    });
  });

  test('typing after the closing tag opens a paragraph after the control, and keeps typing in it', () => {
    const surface = open();
    caretBeside(surface, { text: 'c', offset: 1 }, 'close', 'after');
    surface.type('Z');
    expect(reading(surface)).toBe('a|B[b|c]|Z|d');
    expect(caretText(surface)).toBe('Z:1');
    surface.type('W');
    expect(reading(surface)).toBe('a|B[b|c]|ZW|d');
  });

  test('typing before the opening tag opens a paragraph before the control', () => {
    const surface = open();
    caretBeside(surface, { text: 'b', offset: 0 }, 'open', 'before');
    surface.type('Z');
    expect(reading(surface)).toBe('a|Z|B[b|c]|d');
    expect(caretText(surface)).toBe('Z:1');
  });

  test('Enter after the closing tag opens an empty paragraph after the control, as one undo step', () => {
    const surface = open();
    caretBeside(surface, { text: 'c', offset: 1 }, 'close', 'after');
    surface.splitParagraph();
    expect(reading(surface)).toBe('a|B[b|c]||d');
    expect(caretText(surface)).toBe(':0');
    surface.undo();
    expect(reading(surface)).toBe('a|B[b|c]|d');
  });

  test('pasting there is refused, and nothing lands inside the control', () => {
    const surface = open();
    caretBeside(surface, { text: 'c', offset: 1 }, 'close', 'after');
    surface.insertPlainText('colado');
    expect(reading(surface)).toBe('a|B[b|c]|d');
  });
});

describe('Enter inside a block control', () => {
  test('just inside the closing tag divides the last paragraph within the control', () => {
    const surface = open();
    caretBeside(surface, { text: 'c', offset: 1 }, 'close', 'before');
    surface.splitParagraph();
    expect(reading(surface)).toBe('a|B[b|c|]|d');
  });

  test('with the tags hidden, at the end of the last paragraph, stays within the control', () => {
    const surface = open(false);
    const caret = { paragraphId: paragraphNamed(surface, 'c'), offset: 1 };
    surface.setSelection({ anchor: caret, head: caret });
    surface.splitParagraph();
    expect(reading(surface)).toBe('a|B[b|c|]|d');
  });
});
