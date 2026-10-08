// Backspace and Delete where the caret meets an inline content control. With the tags drawn, a
// key beside a tag selects that tag's control whole and the next one removes it, and a key in a
// control showing its prompt removes it at once; with the tags hidden, a key keeps Word's rule.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import type { OoxmlNode } from '../../store/package/ooxml-tree.ts';
import { contentControlPropertiesOf } from '../../store/package/content-control-nodes.ts';
import {
  contentControlContentChildren,
  isContentControl,
} from '../../store/package/content-control-walk.ts';
import type { PaginatedSurface } from '../paginated-surface.ts';
import { mount, putCaret } from './paginated-surface-fixtures.ts';

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const sdt = (tag: string, inner: string, extra = '') =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/>${extra}<w:richText/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
const prompt = (tag: string) => sdt(tag, run('pp'), '<w:showingPlcHdr/>');

/** `AB O{xyz} CD`: the control spans offsets 2 to 5. */
const OPTIONAL = `${run('AB')}${sdt('O', run('xyz'))}${run('CD')}`;
/** `AB O{pp} CD`: the control shows its prompt. */
const EMPTY = `${run('AB')}${prompt('O')}${run('CD')}`;
/** `A G{B{rg}E{cnh}} Z`. */
const GROUP = `${run('A')}${sdt('G', sdt('B', run('rg')) + sdt('E', run('cnh')))}${run('Z')}`;

const mounted: { surface: PaginatedSurface; container: HTMLElement }[] = [];
afterEach(() => {
  for (const { surface } of mounted.splice(0)) surface.destroy();
});

function plain(body: string): { surface: PaginatedSurface; container: HTMLElement } {
  const result = mount(`<w:p>${body}</w:p>`);
  mounted.push(result);
  return result;
}

function tagged(body: string): PaginatedSurface {
  const { surface } = plain(body);
  surface.setContentControlTags({
    token: 'keys',
    labelsOf: ({ tag }) => ({ open: { text: `${tag}▸` }, close: { text: `◂${tag}` } }),
  });
  return surface;
}

/** The paragraph with each control written `tag{…}`, and `tag*{…}` while it shows its prompt. */
function bracketed(surface: PaginatedSurface): string {
  const paragraph = surface.session.part().root.children[0]!;
  const walk = (nodes: readonly OoxmlNode[]): string =>
    nodes
      .map((node) => {
        if (node.kind === 'textValue') return node.value;
        if (isContentControl(node)) {
          const properties = contentControlPropertiesOf(node);
          const flag = properties.showingPlaceholder ? '*' : '';
          return `${properties.tag}${flag}{${walk(contentControlContentChildren(node))}}`;
        }
        return node.localName === 'sdtPr' || node.localName === 'rPr' ? '' : walk(node.children);
      })
      .join('');
  return walk(paragraph.kind === 'textValue' ? [] : paragraph.children);
}

function controlId(surface: PaginatedSurface, tag: string): string {
  const find = (nodes: readonly OoxmlNode[]): string | null => {
    for (const node of nodes) {
      if (node.kind === 'textValue') continue;
      if (isContentControl(node) && contentControlPropertiesOf(node).tag === tag) return node.id;
      const inner = find(node.children);
      if (inner) return inner;
    }
    return null;
  };
  const id = find([surface.session.part().root]);
  if (id === null) throw new Error(`no control ${tag}`);
  return id;
}

/** The caret at `offset`, in the slot beside the named tag. */
function caretBeside(
  surface: PaginatedSurface,
  offset: number,
  tag: string,
  edge: 'open' | 'close',
  side: 'before' | 'after'
): void {
  const paragraphId = surface.session.paragraphIds()[0]!;
  const caret = { paragraphId, offset };
  surface.setSelection(
    { anchor: caret, head: caret },
    { controlId: controlId(surface, tag), edge, side }
  );
}

const selected = (surface: PaginatedSurface): string => {
  const { anchor, head } = surface.state().selection;
  return `${anchor.offset}-${head.offset}`;
};

describe('with the tags drawn, a key beside a tag', () => {
  test('Backspace behind a closing tag selects the control, and the next one removes it', () => {
    const surface = tagged(OPTIONAL);
    caretBeside(surface, 5, 'O', 'close', 'after');
    surface.deleteBackward();
    expect(bracketed(surface)).toBe('ABO{xyz}CD');
    expect(selected(surface)).toBe('2-5');
    surface.deleteBackward();
    expect(bracketed(surface)).toBe('ABCD');
    expect(selected(surface)).toBe('2-2');
  });

  test('Delete in front of an opening tag selects the control, and the next one removes it', () => {
    const surface = tagged(OPTIONAL);
    caretBeside(surface, 2, 'O', 'open', 'before');
    surface.deleteForward();
    expect(selected(surface)).toBe('2-5');
    surface.deleteForward();
    expect(bracketed(surface)).toBe('ABCD');
  });

  test('Backspace just inside the opening tag never takes the character outside it', () => {
    const surface = tagged(OPTIONAL);
    caretBeside(surface, 2, 'O', 'open', 'after');
    surface.deleteBackward();
    expect(bracketed(surface)).toBe('ABO{xyz}CD');
    expect(selected(surface)).toBe('2-5');
  });

  test('Delete just inside the closing tag never takes the character outside it', () => {
    const surface = tagged(OPTIONAL);
    caretBeside(surface, 5, 'O', 'close', 'before');
    surface.deleteForward();
    expect(bracketed(surface)).toBe('ABO{xyz}CD');
    expect(selected(surface)).toBe('2-5');
  });

  test('the word keys act on the control too', () => {
    const surface = tagged(OPTIONAL);
    caretBeside(surface, 5, 'O', 'close', 'after');
    surface.deleteWordBackward();
    surface.deleteWordBackward();
    expect(bracketed(surface)).toBe('ABCD');
  });

  test('a group goes with every branch in it', () => {
    const surface = tagged(GROUP);
    caretBeside(surface, 6, 'G', 'close', 'after');
    surface.deleteBackward();
    surface.deleteBackward();
    expect(bracketed(surface)).toBe('AZ');
  });

  test('the control selected whole is painted as selected', () => {
    const { surface, container } = plain(OPTIONAL);
    surface.setContentControlTags({
      token: 'keys',
      labelsOf: ({ tag }) => ({ open: { text: `${tag}▸` }, close: { text: `◂${tag}` } }),
    });
    caretBeside(surface, 5, 'O', 'close', 'after');
    surface.deleteBackward();
    const painted = container.querySelector<HTMLElement>(
      '.docx-content-control-chrome[data-selected]'
    );
    expect(painted?.dataset.docxContentControl).toBe(controlId(surface, 'O'));
  });

  test('a wrapper locked against deletion stays', () => {
    const surface = tagged(
      `${run('AB')}${sdt('O', run('xyz'), '<w:lock w:val="sdtLocked"/>')}${run('CD')}`
    );
    caretBeside(surface, 5, 'O', 'close', 'after');
    surface.deleteBackward();
    surface.deleteBackward();
    expect(bracketed(surface)).toBe('ABO{xyz}CD');
  });
});

describe('the control selected whole', () => {
  test('is forgotten when the selection moves away and comes back', () => {
    const surface = tagged(OPTIONAL);
    caretBeside(surface, 5, 'O', 'close', 'after');
    surface.deleteBackward();
    putCaret(surface, 0);
    const paragraphId = surface.session.paragraphIds()[0]!;
    surface.setSelection({
      anchor: { paragraphId, offset: 2 },
      head: { paragraphId, offset: 5 },
    });
    surface.deleteBackward();
    expect(bracketed(surface)).toBe('ABO*{Click here to enter text.}CD');
  });

  test('typing over it replaces what it holds and keeps the control', () => {
    const surface = tagged(OPTIONAL);
    caretBeside(surface, 5, 'O', 'close', 'after');
    surface.deleteBackward();
    surface.type('Q');
    expect(bracketed(surface)).toBe('ABO{Q}CD');
  });

  test('cutting it removes it', () => {
    const surface = tagged(OPTIONAL);
    caretBeside(surface, 5, 'O', 'close', 'after');
    surface.deleteBackward();
    expect(surface.deleteSelection()).toBe(true);
    expect(bracketed(surface)).toBe('ABCD');
  });
});

describe('a key in a control showing its prompt', () => {
  test('Backspace removes the control at once', () => {
    const surface = tagged(EMPTY);
    caretBeside(surface, 2, 'O', 'open', 'after');
    surface.deleteBackward();
    expect(bracketed(surface)).toBe('ABCD');
    expect(selected(surface)).toBe('2-2');
  });

  test('Delete removes the control at once', () => {
    const surface = tagged(EMPTY);
    caretBeside(surface, 2, 'O', 'open', 'after');
    surface.deleteForward();
    expect(bracketed(surface)).toBe('ABCD');
  });

  test('a key over the selected prompt removes the control', () => {
    const surface = tagged(EMPTY);
    const paragraphId = surface.session.paragraphIds()[0]!;
    surface.setSelection({ anchor: { paragraphId, offset: 2 }, head: { paragraphId, offset: 4 } });
    surface.deleteBackward();
    expect(bracketed(surface)).toBe('ABCD');
  });

  test('beside the prompt, outside it, the first key selects it', () => {
    const surface = tagged(EMPTY);
    caretBeside(surface, 4, 'O', 'close', 'after');
    surface.deleteBackward();
    expect(bracketed(surface)).toBe('ABO*{pp}CD');
    expect(selected(surface)).toBe('2-4');
  });
});

describe('with the tags hidden', () => {
  test('a key in a prompt keeps the Word rule and leaves the control', () => {
    const { surface } = plain(EMPTY);
    putCaret(surface, 2);
    surface.deleteBackward();
    expect(bracketed(surface)).toBe('AO*{pp}CD');
  });

  test('Backspace at a control edge takes one character, as in Word', () => {
    const { surface } = plain(OPTIONAL);
    putCaret(surface, 5);
    surface.deleteBackward();
    expect(bracketed(surface)).toBe('ABO{xy}CD');
  });
});
