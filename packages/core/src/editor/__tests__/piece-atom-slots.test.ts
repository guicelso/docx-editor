// An atom is one unit with its tags: the caret meets it before or after it, never inside. A tag the
// host draws on an atom stands for the whole atom, so typing beside it lands outside, a key at its
// outer edge takes it whole with or without tags, and its tags carry the atom's selection.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import type {
  ContentControlTagLabel,
  PieceAtom,
} from '../../contracts/editor-content-control-view.ts';
import type { ContentControlView } from '../../layout/content-control-view.ts';
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

/** Offsets: `A ` is 0–2, `Qual` 2–6, ` fim` 6–10. */
const FILLED = `<w:p>${run('A ')}${sdt('q', run('Qual'))}${run(' fim')}</w:p>`;
/** Offsets: `A ` is 0–2, the empty place stands at 2, `fim` 2–5. */
const EMPTY = `<w:p>${run('A ')}${sdt('p', '')}${run('fim')}</w:p>`;

const ATOM: PieceAtom = { tone: 'citation', form: 'atom' };

const mounted: PaginatedSurface[] = [];
afterEach(() => {
  for (const surface of mounted.splice(0)) surface.destroy();
});

type Labels = { readonly open?: ContentControlTagLabel; readonly close?: ContentControlTagLabel };

/** A body whose one control the host makes an atom, drawn with `labels` (none: no tags at all). */
function atomOf(
  body: string,
  labels: Labels | null
): { surface: PaginatedSurface; container: HTMLElement } {
  const opened = mount(body);
  mounted.push(opened.surface);
  const view: ContentControlView = {
    pieces: { pieceOf: () => ({ atom: ATOM }) },
    ...(labels === null ? {} : { tags: { labelsOf: () => labels } }),
  };
  opened.surface.setContentControlView(view);
  return opened;
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

describe('an atom with a tag at its close', () => {
  const CONDITION: Labels = { close: { text: 'se houver', tone: 'optional' } };

  test('arriving from the left, the caret stands after the tag, and typing lands outside', () => {
    const { surface } = atomOf(FILLED, CONDITION);
    putCaret(surface, 2);
    surface.navigate('right');
    expect(surface.state().selection.head.offset).toBe(6);
    surface.type('Z');
    expect(bracketed(surface)).toBe('A q{Qual}Z fim');
  });

  test('leaving to the left steps over the atom in one step', () => {
    const { surface } = atomOf(FILLED, CONDITION);
    putCaret(surface, 2);
    surface.navigate('right');
    surface.navigate('left');
    expect(surface.state().selection.head.offset).toBe(2);
    surface.type('Y');
    expect(bracketed(surface)).toBe('A Yq{Qual} fim');
  });
});

describe('an empty atom drawn by its opening tag', () => {
  const PLACE: Labels = { open: { text: 'Cônjuge', tone: 'alternatives', mark: 'nested-part' } };

  test('the slot after the tag stands after the atom: typing there lands outside', () => {
    const { surface } = atomOf(EMPTY, PLACE);
    putCaret(surface, 2);
    const slot = surface.state().contentControls.caretSlot;
    expect(slot?.placement).toEqual({
      beside: { controlId: slot!.left!.controlId, side: 'after' },
    });
    surface.type('Z');
    expect(bracketed(surface)).toBe('A p{}Zfim');
  });

  test('one step to the left stands before it', () => {
    const { surface } = atomOf(EMPTY, PLACE);
    putCaret(surface, 2);
    surface.navigate('left');
    surface.type('Y');
    expect(bracketed(surface)).toBe('A Yp{}fim');
  });

  test('Backspace after it selects it whole, its tag included, and the next removes it', () => {
    const { surface, container } = atomOf(EMPTY, PLACE);
    putCaret(surface, 2);
    surface.deleteBackward();
    const tag = container.querySelector<HTMLElement>('.docx-cc-tag[data-piece-control]');
    expect(tag?.dataset.pieceControl).toBe(tag?.dataset.ccTagControl);
    expect(tag?.getAttribute('data-selected')).toBe('true');
    surface.deleteBackward();
    expect(bracketed(surface)).toBe('A fim');
  });
});

describe('an atom with no tags', () => {
  test('Backspace at its end selects it whole, and the next removes it', () => {
    const { surface } = atomOf(FILLED, null);
    putCaret(surface, 6);
    surface.deleteBackward();
    const { anchor, head } = surface.state().selection;
    expect([anchor.offset, head.offset]).toEqual([2, 6]);
    expect(bracketed(surface)).toBe('A q{Qual} fim');
    surface.deleteBackward();
    expect(bracketed(surface)).toBe('A  fim');
  });

  test('Delete at its start selects it whole', () => {
    const { surface } = atomOf(FILLED, null);
    putCaret(surface, 2);
    surface.deleteForward();
    const { anchor, head } = surface.state().selection;
    expect([anchor.offset, head.offset]).toEqual([2, 6]);
  });
});

describe('the tags of a control that is no atom', () => {
  test('name no control, so the atom state never reaches them', () => {
    const opened = mount(FILLED);
    mounted.push(opened.surface);
    opened.surface.setContentControlView({
      tags: { labelsOf: () => ({ open: { text: 'Trecho ▸' } }) },
    });
    const tag = opened.container.querySelector<HTMLElement>('.docx-cc-tag');
    expect(tag?.dataset.ccTagControl).toBeDefined();
    expect(tag?.hasAttribute('data-piece-control')).toBe(false);
  });
});

describe('the role of a piece', () => {
  test('lands on its chip and on its frame as data-piece-role', () => {
    const opened = mount(FILLED);
    mounted.push(opened.surface);
    opened.surface.setContentControlView({
      tags: {
        labelsOf: () => ({ open: { text: 'Grupo ▸', tone: 'alternatives', role: 'container' } }),
      },
      pieces: { pieceOf: () => ({ frame: { tone: 'alternatives', role: 'container' } }) },
    });
    opened.surface.setSelection({
      anchor: { paragraphId: opened.surface.session.paragraphIds()[0]!, offset: 3 },
      head: { paragraphId: opened.surface.session.paragraphIds()[0]!, offset: 3 },
    });
    const tag = opened.container.querySelector<HTMLElement>('.docx-cc-tag');
    expect(tag?.dataset.pieceRole).toBe('container');
    const chrome = opened.container.querySelector<HTMLElement>('[data-docx-content-control]');
    expect(chrome?.dataset.pieceRole).toBe('container');
  });

  test('outside the pattern is not published', () => {
    const opened = mount(FILLED);
    mounted.push(opened.surface);
    opened.surface.setContentControlView({
      tags: { labelsOf: () => ({ open: { text: 'Grupo ▸', role: 'no spaces' } }) },
    });
    const tag = opened.container.querySelector<HTMLElement>('.docx-cc-tag');
    expect(tag?.hasAttribute('data-piece-role')).toBe(false);
  });
});
