// The caret slot on a mounted surface: which slot the caret stands in at a tagged edge, how the
// arrows walk the slots, and where typing in each one lands.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import type { OoxmlNode } from '../../store/package/ooxml-tree.ts';
import {
  contentControlContentChildren,
  isContentControl,
} from '../../store/package/content-control-walk.ts';
import { contentControlTagSubjectOf } from '../../layout/content-control-tags.ts';
import type { PaginatedSurface } from '../paginated-surface.ts';
import { createPublishSignal } from '../surface-publish-signal.ts';
import { mount, putCaret } from './paginated-surface-fixtures.ts';

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const sdt = (tag: string, inner: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

/** `CPF G{B{RG}E{CNH}}`, ending the paragraph — the shape of a block body. */
const GROUP = `${run('CPF ')}${sdt('G', sdt('B', run('RG')) + sdt('E', run('CNH')))}`;

const mounted: PaginatedSurface[] = [];
afterEach(() => {
  for (const surface of mounted.splice(0)) surface.destroy();
});

function tagged(body: string): PaginatedSurface {
  const { surface } = mount(`<w:p>${body}</w:p>`);
  mounted.push(surface);
  surface.setContentControlTags({
    token: 'slots',
    labelsOf: ({ tag }) => ({ open: { text: `${tag}▸` }, close: { text: `◂${tag}` } }),
  });
  return surface;
}

/** The paragraph with its controls bracketed by tag: `CPF G{B{RG}E{CNH}}`. */
function bracketed(surface: PaginatedSurface): string {
  const paragraph = surface.session.part().root.children[0]!;
  const walk = (nodes: readonly OoxmlNode[]): string =>
    nodes
      .map((node) => {
        if (node.kind === 'textValue') return node.value;
        if (isContentControl(node)) {
          return `${contentControlTagSubjectOf(node).tag}{${walk(contentControlContentChildren(node))}}`;
        }
        return node.localName === 'sdtPr' || node.localName === 'rPr' ? '' : walk(node.children);
      })
      .join('');
  return walk(paragraph.kind === 'textValue' ? [] : paragraph.children);
}

const slotOf = (surface: PaginatedSurface) => {
  const slot = surface.state().contentControls.caretSlot;
  const name = (side: { controlId: string; edge: string } | null) =>
    side === null ? '|' : `${side.edge}`;
  return slot === null ? null : `${name(slot.left)} ${name(slot.right)}`;
};

describe('the caret slot at a tagged edge', () => {
  test('a programmatic caret stands in the slot touching the text on the right', () => {
    const surface = tagged(GROUP);
    putCaret(surface, 4);
    expect(slotOf(surface)).toBe('open |');
    surface.type('X');
    expect(bracketed(surface)).toBe('CPF G{B{XRG}E{CNH}}');
  });

  test('the left arrow walks the slots of the offset before it moves it', () => {
    const surface = tagged(GROUP);
    putCaret(surface, 4);
    surface.navigate('left');
    expect(slotOf(surface)).toBe('open open');
    surface.navigate('left');
    expect(slotOf(surface)).toBe('| open');
    surface.navigate('left');
    expect(surface.state().selection.head.offset).toBe(3);
  });

  test('arriving from the left stops in the slot touching the text it left', () => {
    const surface = tagged(GROUP);
    putCaret(surface, 3);
    surface.navigate('right');
    expect(surface.state().selection.head.offset).toBe(4);
    expect(slotOf(surface)).toBe('| open');
  });

  test('in front of the group, typing stays in front of it, keystroke after keystroke', () => {
    const surface = tagged(GROUP);
    putCaret(surface, 4);
    surface.navigate('left');
    surface.navigate('left');
    surface.type('A');
    surface.type('B');
    expect(bracketed(surface)).toBe('CPF ABG{B{RG}E{CNH}}');
  });

  test('between two branches, typing lands behind the first, inside the group', () => {
    const surface = tagged(GROUP);
    putCaret(surface, 6);
    surface.navigate('left');
    expect(slotOf(surface)).toBe('close open');
    surface.type('X');
    expect(bracketed(surface)).toBe('CPF G{B{RG}XE{CNH}}');
  });

  test('at the end the paragraph shares with the group, typing lands outside every control', () => {
    const surface = tagged(GROUP);
    putCaret(surface, 9);
    expect(slotOf(surface)).toBe('close |');
    surface.type('.');
    expect(bracketed(surface)).toBe('CPF G{B{RG}E{CNH}}.');
  });

  test('a shift arrow extends by offset, leaving the slot behind', () => {
    const surface = tagged(GROUP);
    putCaret(surface, 4);
    surface.navigate('left');
    surface.navigate('left', true);
    const { anchor, head } = surface.state().selection;
    expect([anchor.offset, head.offset]).toEqual([4, 3]);
    expect(surface.state().contentControls.caretSlot).toBeNull();
  });

  test('backspace deletes by offset, whichever slot the caret stands in', () => {
    const surface = tagged(GROUP);
    putCaret(surface, 4);
    surface.navigate('left');
    surface.deleteBackward();
    expect(bracketed(surface)).toBe('CPFG{B{RG}E{CNH}}');
  });

  test('a host puts the caret in a chosen slot, as a press on the tag would', () => {
    const surface = tagged(GROUP);
    putCaret(surface, 4);
    surface.navigate('left');
    surface.navigate('left');
    const group = surface.state().contentControls.caretSlot!.right!.controlId;
    putCaret(surface, 9);
    const paragraphId = surface.state().selection.head.paragraphId;
    const at = { paragraphId, offset: 4 };
    surface.setSelection(
      { anchor: at, head: at },
      { controlId: group, edge: 'open', side: 'after' }
    );
    expect(slotOf(surface)).toBe('open open');
    surface.type('X');
    expect(bracketed(surface)).toBe('CPF G{XB{RG}E{CNH}}');
  });

  test('a step between the slots of one offset is a publish the host hears', () => {
    const surface = tagged(GROUP);
    putCaret(surface, 4);
    const signal = createPublishSignal();
    signal.adopt(surface);
    surface.navigate('left');
    expect(surface.state().selection.head.offset).toBe(4);
    expect(signal.moved(surface.state(), surface)).toBe(true);
    expect(signal.moved(surface.state(), surface)).toBe(false);
  });

  test('without tags there is no slot, and typing keeps the store’s own rule', () => {
    const surface = tagged(GROUP);
    surface.setContentControlTags(null);
    putCaret(surface, 4);
    expect(surface.state().contentControls.caretSlot).toBeNull();
  });
});
