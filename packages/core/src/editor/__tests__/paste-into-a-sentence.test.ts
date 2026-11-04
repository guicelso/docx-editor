// A paste lands where the caret names, as typing does, and inside an inline control it is read as
// one sentence: paragraphs and lines join with a space, since a paragraph mark has no place there.
// A copy carries its last paragraph mark only with the whole paragraph, and the host's paste
// policy rewrites what lands and may refuse what the paste would publish.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { mount, paragraph, putCaret } from './paginated-surface-fixtures.ts';

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const control = (tag: string, inner: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w:richText/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

/** "a " then the control S holding "xy": offset 3 stands between x and y, inside S. */
const SENTENCE = `<w:p>${run('a ')}${control('S', run('xy'))}</w:p>`;

function twoParagraphsCopied(): { text: string; html: string | null } {
  const source = mount(paragraph('bold') + paragraph('plain'));
  putCaret(source.surface, 0);
  source.surface.selectAll();
  return source.surface.copyFlavours();
}

function select(surface: ReturnType<typeof mount>['surface'], start: number, end: number): void {
  const paragraphId = surface.session.paragraphIds()[0]!;
  surface.setSelection({
    anchor: { paragraphId, offset: start },
    head: { paragraphId, offset: end },
  });
}

describe('the paragraph mark a copy carries', () => {
  test('a range from mid-paragraph to its end carries text, the whole paragraph its mark', () => {
    const { surface } = mount(paragraph('first second'));
    select(surface, 6, 12);
    expect(surface.copyFlavours().html).toContain('data-docx-fragment-end="open"');
    select(surface, 0, 12);
    expect(surface.copyFlavours().html).toContain('data-docx-fragment-end="covered"');
  });

  test('the end of a paragraph pasted mid-paragraph does not divide the host', () => {
    const source = mount(paragraph('first second'));
    select(source.surface, 6, 12);
    const copied = source.surface.copyFlavours();
    const target = mount(paragraph('ab'));
    putCaret(target.surface, 1);
    target.surface.pasteRich(copied.text, copied.html);
    expect(target.surface.session.paragraphIds()).toHaveLength(1);
    expect(target.surface.session.bodyText()).toBe('asecondb');
  });
});

describe('a paste inside an inline control', () => {
  test('a caret between its characters names it, and paragraphs join into its sentence', () => {
    const copied = twoParagraphsCopied();
    const { surface } = mount(SENTENCE);
    putCaret(surface, 3);
    surface.pasteRich(copied.text, copied.html);
    expect(surface.session.paragraphIds()).toHaveLength(1);
    expect(surface.session.bodyText()).toBe('a xbold plainy');
    expect(serializeOoxmlPart(surface.session.part())).toMatch(/<w:sdtContent>.*bold.*plain.*<\/w:sdtContent>/);
  });

  test('a selection inside it is replaced in it, as one sentence', () => {
    const copied = twoParagraphsCopied();
    const { surface } = mount(`<w:p>${run('a ')}${control('S', run('xyz'))}</w:p>`);
    select(surface, 3, 4);
    surface.pasteRich(copied.text, copied.html);
    expect(surface.session.bodyText()).toBe('a xbold plainz');
    expect(surface.session.paragraphIds()).toHaveLength(1);
  });

  test('plain text lines join with a space', () => {
    const { surface } = mount(SENTENCE);
    putCaret(surface, 3);
    surface.insertPlainText('one\n\n  two  \n');
    expect(surface.session.paragraphIds()).toHaveLength(1);
    expect(surface.session.bodyText()).toBe('a xone twoy');
  });

  test('outside every control, plain lines still divide the paragraph', () => {
    const { surface } = mount(paragraph('ab'));
    putCaret(surface, 1);
    surface.insertPlainText('one\ntwo');
    expect(surface.session.paragraphIds()).toHaveLength(2);
  });
});

describe("the host's paste policy", () => {
  test('the blocks it answers land, read against the part as it stands', () => {
    const copied = twoParagraphsCopied();
    const { surface } = mount(paragraph(''));
    const seen: string[] = [];
    surface.setPastePolicy({
      blocksOf: ({ blocks, part }) => {
        seen.push(part.name);
        return blocks.slice(0, 1);
      },
    });
    putCaret(surface, 0);
    surface.pasteRich(copied.text, copied.html);
    expect(seen).toEqual(['/word/document.xml']);
    expect(surface.session.bodyText()).toContain('bold');
    expect(surface.session.bodyText()).not.toContain('plain');
  });

  test('a refusal leaves the document as it was, says why, and no plain text lands', () => {
    const copied = twoParagraphsCopied();
    const { surface } = mount(paragraph('ab'));
    surface.setPastePolicy({
      admit: ({ before, after }) => (before === after ? null : 'not here'),
    });
    putCaret(surface, 1);
    surface.pasteRich(copied.text, copied.html);
    expect(surface.session.bodyText()).toBe('ab');
    expect(surface.state().lastRejection).toBe('not here');
  });
});
