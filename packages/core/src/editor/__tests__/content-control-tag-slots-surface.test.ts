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
/** A control showing its prompt, `pp`: the first key typed into it replaces the whole prompt. */
const prompt = (tag: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w:showingPlcHdr/></w:sdtPr><w:sdtContent>${run('pp')}</w:sdtContent></w:sdt>`;

/** `CPF G{B{RG}E{CNH}}`, ending the paragraph — the shape of a block body. */
const GROUP = `${run('CPF ')}${sdt('G', sdt('B', run('RG')) + sdt('E', run('CNH')))}`;
/** `CPF G{B{pp}E{pp}}`: the branches a variation creates, side by side and empty. */
const EMPTY_BRANCHES = `${run('CPF ')}${sdt('G', prompt('B') + prompt('E'))}`;

const mounted: PaginatedSurface[] = [];
afterEach(() => {
  for (const surface of mounted.splice(0)) surface.destroy();
});

function plain(body: string): PaginatedSurface {
  const { surface } = mount(`<w:p>${body}</w:p>`);
  mounted.push(surface);
  return surface;
}

function tagged(body: string): PaginatedSurface {
  const surface = plain(body);
  surface.setContentControlView({
    tags: { labelsOf: ({ tag }) => ({ open: { text: `${tag}▸` }, close: { text: `◂${tag}` } }) },
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

function controlIdOf(surface: PaginatedSurface, tag: string): string {
  const find = (nodes: readonly OoxmlNode[]): string | null => {
    for (const node of nodes) {
      if (node.kind === 'textValue') continue;
      if (isContentControl(node) && contentControlTagSubjectOf(node).tag === tag) return node.id;
      const inner = find(node.children);
      if (inner) return inner;
    }
    return null;
  };
  const id = find([surface.session.part().root]);
  if (id === null) throw new Error(`no control ${tag}`);
  return id;
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

  test('backspace beside a tag selects its control instead of deleting across the tag', () => {
    const surface = tagged(GROUP);
    putCaret(surface, 4);
    surface.navigate('left');
    surface.deleteBackward();
    expect(bracketed(surface)).toBe('CPF G{B{RG}E{CNH}}');
    const { anchor, head } = surface.state().selection;
    expect([anchor.offset, head.offset]).toEqual([4, 9]);
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
    surface.setContentControlView(null);
    putCaret(surface, 4);
    expect(surface.state().contentControls.caretSlot).toBeNull();
  });
});

describe('the caret after an edit stays where the edit happened', () => {
  test('typing at the end of a branch stays in it, keystroke after keystroke', () => {
    const surface = tagged(GROUP);
    putCaret(surface, 6);
    surface.navigate('left');
    surface.navigate('left');
    expect(slotOf(surface)).toBe('| close');
    surface.type('X');
    surface.type('Y');
    expect(bracketed(surface)).toBe('CPF G{B{RGXY}E{CNH}}');
    expect(slotOf(surface)).toBe('| close');
  });

  test('typing into an empty branch replaces its prompt and stays in it', () => {
    const surface = tagged(EMPTY_BRANCHES);
    putCaret(surface, 5);
    surface.type('d');
    surface.type('e');
    expect(bracketed(surface)).toBe('CPF G{B{de}E{pp}}');
  });

  test('in front of a group whose first branch is empty, the text lands in front', () => {
    const surface = tagged(EMPTY_BRANCHES);
    putCaret(surface, 4);
    surface.navigate('left');
    surface.navigate('left');
    expect(slotOf(surface)).toBe('| open');
    surface.type('X');
    expect(bracketed(surface)).toBe('CPF XG{B{pp}E{pp}}');
  });

  test('Backspace takes the character and keeps what stood on the caret’s right', () => {
    const surface = tagged(GROUP);
    putCaret(surface, 6);
    surface.navigate('left');
    surface.navigate('left');
    surface.deleteBackward();
    surface.type('X');
    expect(bracketed(surface)).toBe('CPF G{B{RX}E{CNH}}');
  });

  test('Delete takes the character and keeps what stood on the caret’s left', () => {
    const surface = tagged(GROUP);
    putCaret(surface, 5);
    surface.deleteForward();
    surface.type('X');
    expect(bracketed(surface)).toBe('CPF G{B{RX}E{CNH}}');
  });

  test('deleting a selection leaves the caret where the selection began', () => {
    const surface = tagged(GROUP);
    const paragraphId = surface.state().selection.head.paragraphId;
    surface.setSelection({
      anchor: { paragraphId, offset: 5 },
      head: { paragraphId, offset: 6 },
    });
    surface.deleteSelection();
    surface.type('X');
    expect(bracketed(surface)).toBe('CPF G{B{RX}E{CNH}}');
  });

  test('a plain paste lands in the slot, and typing continues after it', () => {
    const surface = tagged(GROUP);
    putCaret(surface, 6);
    surface.navigate('left');
    expect(slotOf(surface)).toBe('close open');
    surface.insertPlainText('X');
    surface.type('Y');
    expect(bracketed(surface)).toBe('CPF G{B{RG}XYE{CNH}}');
  });

  test('a rich paste lands in the slot, as a plain one does', () => {
    const surface = tagged(GROUP);
    putCaret(surface, 4);
    surface.navigate('left');
    expect(slotOf(surface)).toBe('open open');
    expect(surface.pasteRich('X', '<b>X</b>')).toBe(true);
    expect(bracketed(surface)).toBe('CPF G{XB{RG}E{CNH}}');
  });

  test('a composed accent lands in the slot, as a typed one does', () => {
    const { surface, container } = mount(`<w:p>${GROUP}</w:p>`);
    mounted.push(surface);
    surface.setContentControlView({
      tags: { labelsOf: ({ tag }) => ({ open: { text: `${tag}▸` }, close: { text: `◂${tag}` } }) },
    });
    putCaret(surface, 6);
    surface.navigate('left');
    expect(slotOf(surface)).toBe('close open');
    const pages = container.querySelector('.docx-pages')!;
    pages.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }));
    const branch = [
      ...container.querySelectorAll<HTMLElement>('[data-paragraph-id][data-start]'),
    ].find((span) => span.textContent === 'RG')!;
    branch.textContent = 'RGé';
    const end = new CompositionEvent('compositionend', { bubbles: true });
    // happy-dom drops `data` from the init dict, which is where a browser puts the composed string.
    Object.defineProperty(end, 'data', { value: 'é' });
    pages.dispatchEvent(end);
    expect(bracketed(surface)).toBe('CPF G{B{RG}éE{CNH}}');
  });
});

describe('a host that wrote text puts the caret after it', () => {
  test('the caret touches what the host wrote, and typing continues there', () => {
    const surface = tagged(GROUP);
    const paragraphId = surface.state().selection.head.paragraphId;
    const branch = controlIdOf(surface, 'B');
    surface.applyAutomationOps(() => [
      { op: 'insertText', paragraphId, offset: 6, text: 'X', inside: branch },
    ]);
    const after = { paragraphId, offset: 7 };
    surface.setSelection({ anchor: after, head: after }, { afterText: true });
    expect(slotOf(surface)).toBe('| close');
    surface.type('Y');
    expect(bracketed(surface)).toBe('CPF G{B{RGXY}E{CNH}}');
  });
});

describe('without tags, an edit still leaves the caret where it happened', () => {
  test('typing into an empty branch beside another stays in it', () => {
    const surface = plain(EMPTY_BRANCHES);
    putCaret(surface, 5);
    surface.type('d');
    surface.type('e');
    expect(bracketed(surface)).toBe('CPF G{B{de}E{pp}}');
  });

  test('Delete before the last character of a branch keeps the caret in it', () => {
    const surface = plain(GROUP);
    putCaret(surface, 5);
    surface.deleteForward();
    surface.type('X');
    expect(bracketed(surface)).toBe('CPF G{B{RX}E{CNH}}');
  });

  test('the slot an edit placed is published, and a host can place the caret in one', () => {
    const surface = plain(GROUP);
    putCaret(surface, 5);
    surface.deleteForward();
    expect(slotOf(surface)).toBe('| close');
    const at = { paragraphId: surface.state().selection.head.paragraphId, offset: 4 };
    surface.setSelection(
      { anchor: at, head: at },
      { controlId: controlIdOf(surface, 'G'), edge: 'open', side: 'after' }
    );
    expect(slotOf(surface)).toBe('open open');
    surface.type('Y');
    expect(bracketed(surface)).toBe('CPF G{YB{R}E{CNH}}');
  });

  test('a caret no edit placed keeps the Word rule: the control opening there takes the key', () => {
    const surface = plain(GROUP);
    putCaret(surface, 6);
    surface.type('X');
    surface.type('Y');
    expect(bracketed(surface)).toBe('CPF G{B{RG}E{XYCNH}}');
    expect(surface.state().contentControls.caretSlot).toBeNull();
  });
});
