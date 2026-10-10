// An edge tag gives a control's side a slot of its own where nothing is drawn. In a list
// `L{A{o RG}B{a CNH}}` whose first item closes with ", " and whose items draw nothing where they
// open, typing right after the separator lands in the second item only because its open side is
// an edge: without it, the slot after the separator is beside the first item, in the list.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import type { ContentControlTagLabel } from '../../contracts/editor-content-control-view.ts';
import { contentControlSubjectOf } from '../../layout/content-control-properties.ts';
import {
  contentControlContentChildren,
  isContentControl,
} from '../../store/package/content-control-walk.ts';
import type { OoxmlNode } from '../../store/package/ooxml-tree.ts';
import type { PaginatedSurface } from '../paginated-surface.ts';
import { mount, putCaret } from './paginated-surface-fixtures.ts';

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const sdt = (tag: string, inner: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

/** `Doc L{A{o RG}B{a CNH}}`: the list ends the paragraph, as a list ends a block body. */
const LIST = `${run('Doc ')}${sdt('L', sdt('A', run('o RG')) + sdt('B', run('a CNH')))}`;

const mounted: PaginatedSurface[] = [];
afterEach(() => {
  for (const surface of mounted.splice(0)) surface.destroy();
});

type Labels = { readonly open?: ContentControlTagLabel; readonly close?: ContentControlTagLabel };

const EDGE: ContentControlTagLabel = { variant: 'edge' };

/** The list with its tags; `edges` decides whether an item's undrawn sides are edge tags. */
function listed(edges: boolean): PaginatedSurface {
  const { surface } = mount(`<w:p>${LIST}</w:p>`);
  mounted.push(surface);
  const edge = edges ? { open: EDGE } : {};
  const labels: Record<string, Labels> = {
    L: { open: { text: 'L▸' }, close: { text: '◂' } },
    A: { ...edge, close: { text: ', ', variant: 'text' } },
    B: { ...edge, ...(edges ? { close: EDGE } : {}) },
  };
  surface.setContentControlView({
    tags: { labelsOf: ({ tag }) => (tag === undefined ? null : (labels[tag] ?? null)) },
  });
  return surface;
}

/** The paragraph with its controls bracketed by tag. */
function bracketed(surface: PaginatedSurface): string {
  const paragraph = surface.session.part().root.children[0]!;
  const walk = (nodes: readonly OoxmlNode[]): string =>
    nodes
      .map((node) => {
        if (node.kind === 'textValue') return node.value;
        if (isContentControl(node)) {
          return `${contentControlSubjectOf(node).tag}{${walk(contentControlContentChildren(node))}}`;
        }
        return node.localName === 'sdtPr' || node.localName === 'rPr' ? '' : walk(node.children);
      })
      .join('');
  return walk(paragraph.kind === 'textValue' ? [] : paragraph.children);
}

/** Offsets: `Doc ` is 0–4, `o RG` 4–8, `a CNH` 8–13. */
describe('an edge tag on a mounted surface', () => {
  test('typing at the start of an item after a separator lands in that item', () => {
    const surface = listed(true);
    putCaret(surface, 9);
    surface.navigate('left');
    surface.type('X');
    expect(bracketed(surface)).toBe('Doc L{A{o RG}B{Xa CNH}}');
  });

  test('without the edge, the same keystroke falls between the two items', () => {
    const surface = listed(false);
    putCaret(surface, 9);
    surface.navigate('left');
    surface.type('X');
    expect(bracketed(surface)).toBe('Doc L{A{o RG}XB{a CNH}}');
  });

  test('typing at the end of the last item stays in the item', () => {
    const surface = listed(true);
    putCaret(surface, 12);
    surface.navigate('right');
    surface.type('S');
    expect(bracketed(surface)).toBe('Doc L{A{o RG}B{a CNHS}}');
  });

  test('typing at the start of the first item lands in it', () => {
    const surface = listed(true);
    putCaret(surface, 5);
    surface.navigate('left');
    surface.type('Y');
    expect(bracketed(surface)).toBe('Doc L{A{Yo RG}B{a CNH}}');
  });
});
