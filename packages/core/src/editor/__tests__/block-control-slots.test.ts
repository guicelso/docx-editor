// The caret slots beside the tags of a BLOCK-level control: the slot inside its opening tag is the
// start of its first paragraph, and every slot names the level of the control on each side.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import type { OoxmlNode } from '../../store/package/ooxml-tree.ts';
import { contentControlPropertiesOf } from '../../store/package/content-control-nodes.ts';
import type { PaginatedSurface } from '../paginated-surface.ts';
import { mount } from './paginated-surface-fixtures.ts';
import { slotPlacementOf } from '../surface-caret-slots.ts';

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const paragraph = (inner: string) => `<w:p>${inner}</w:p>`;
const control = (tag: string, inner: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w:richText/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

/** `a | B[ I{x}y ]`: the block opens where the inline control opens. */
const BODY = paragraph(run('a')) + control('B', paragraph(control('I', run('x')) + run('y')));

const mounted: PaginatedSurface[] = [];
afterEach(() => {
  for (const surface of mounted.splice(0)) surface.destroy();
});

function tagged(): PaginatedSurface {
  const { surface } = mount(BODY);
  mounted.push(surface);
  surface.setContentControlView({
    tags: { labelsOf: ({ tag }) => ({ open: { text: `${tag}▸` }, close: { text: `◂${tag}` } }) },
  });
  return surface;
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

function nodeOf(surface: PaginatedSurface, id: string): OoxmlNode {
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

const tagged_ = (surface: PaginatedSurface, tag: string) =>
  idOf(
    surface,
    (node) => node.kind === 'contentControl' && contentControlPropertiesOf(node).tag === tag
  );

function textOf(node: OoxmlNode): string {
  if (node.kind === 'textValue') return node.value;
  return node.children.map(textOf).join('');
}

/** The caret at the start of the block's paragraph, in the slot beside the named tag. */
function caretBeside(surface: PaginatedSurface, tag: string, side: 'before' | 'after'): void {
  const paragraphId = idOf(surface, (node) => node.kind === 'paragraph' && textOf(node) === 'xy');
  const caret = { paragraphId, offset: 0 };
  surface.setSelection(
    { anchor: caret, head: caret },
    { controlId: tagged_(surface, tag), edge: 'open', side }
  );
}

describe('the slots beside a block tag', () => {
  test('inside the opening tag is the start of the first paragraph, outside the inline one', () => {
    const surface = tagged();
    caretBeside(surface, 'B', 'after');
    surface.type('Z');
    const inline = nodeOf(surface, tagged_(surface, 'I'));

    expect(textOf(nodeOf(surface, tagged_(surface, 'B')))).toBe('Zxy');
    expect(textOf(inline)).toBe('x');
  });

  test('inside the opening tag, with text right after it, is the start of the first paragraph', () => {
    const { surface } = mount(paragraph(run('a')) + control('C', paragraph(run('w'))));
    mounted.push(surface);
    surface.setContentControlView({
      tags: { labelsOf: ({ tag }) => ({ open: { text: `${tag}▸` }, close: { text: `◂${tag}` } }) },
    });
    const paragraphId = idOf(surface, (node) => node.kind === 'paragraph' && textOf(node) === 'w');
    const caret = { paragraphId, offset: 0 };
    surface.setSelection(
      { anchor: caret, head: caret },
      { controlId: tagged_(surface, 'C'), edge: 'open', side: 'after' }
    );
    surface.type('Z');

    expect(textOf(nodeOf(surface, tagged_(surface, 'C')))).toBe('Zw');
  });

  test('the slot the host reads names the level of the control on each side', () => {
    const surface = tagged();
    caretBeside(surface, 'I', 'before');

    expect(surface.state().contentControls.caretSlot).toEqual({
      left: { controlId: tagged_(surface, 'B'), edge: 'open', level: 'block' },
      right: { controlId: tagged_(surface, 'I'), edge: 'open', level: 'inline' },
    });
  });

  test('a slot inside a block tag names no destination: the offset is already inside it', () => {
    const block = { controlId: 'B', edge: 'open', level: 'block' } as const;
    const inline = { controlId: 'I', edge: 'open', level: 'inline' } as const;

    expect(slotPlacementOf({ left: block, right: null })).toBeNull();
    expect(slotPlacementOf({ left: block, right: inline })).toEqual({
      beside: { controlId: 'I', side: 'before' },
    });
    expect(
      slotPlacementOf({ left: { ...inline, edge: 'close' }, right: { ...block, edge: 'close' } })
    ).toEqual({
      beside: { controlId: 'I', side: 'after' },
    });
  });
});
