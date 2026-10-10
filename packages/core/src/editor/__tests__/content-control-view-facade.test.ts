// The host's view lives on the editor: every surface the editor mounts takes it, so a load, a
// detach and an attach keep the tags, the prompts, the field tones and the field selection.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import type { DocxEditorInstance } from '../docx-editor-types.ts';
import { docx } from './paginated-surface-fixtures.ts';
import { mountAnchorEditor } from './scroll-to-anchor-fixture.ts';

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const control = (tag: string, inner: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w:richText/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
const field = (instruction: string, result: string) =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> MERGEFIELD "${instruction}" </w:instrText></w:r>` +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${result}</w:t></w:r>` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

const placeholder = (tag: string, text: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w:showingPlcHdr/></w:sdtPr><w:sdtContent>${run(text)}</w:sdtContent></w:sdt>`;

const FIRST = docx(`<w:p>${run('um ')}${control('a', run('A'))}${field('field:x', '«x»')}</w:p>`);
const SECOND = docx(
  `<w:p>${run('dois ')}${control('b', run('B'))}${field('field:y', '«y»')}</w:p>`
);
const PROMPTED = docx(`<w:p>${run('tres ')}${placeholder('p', 'Clique aqui')}</w:p>`);

const cleanups: (() => void)[] = [];
afterEach(() => {
  for (const cleanup of cleanups.splice(0).reverse()) cleanup();
});

function mounted(): { editor: DocxEditorInstance; host: HTMLElement } {
  const opened = mountAnchorEditor(FIRST);
  cleanups.push(opened.destroy);
  return opened;
}

const chipsIn = (host: HTMLElement) =>
  [...host.querySelectorAll<HTMLElement>('[data-cc-tag-control]')].map(
    (chip) => chip.textContent?.replace(/[\u00A0\u202F]/g, ' ').trim() ?? ''
  );

const paintedText = (host: HTMLElement) =>
  host.querySelector('.docx-pages')?.textContent?.replace(/[\u00A0\u202F]/g, ' ') ?? '';

const tonesIn = (host: HTMLElement) =>
  [...host.querySelectorAll<HTMLElement>('[data-field-atom]')].map(
    (atom) => atom.dataset.fieldTone
  );

/** Press the middle of the painted field, as a reader would. */
function pressField(editor: DocxEditorInstance, host: HTMLElement): void {
  const surface = editor.surface!;
  const pages = host.querySelector<HTMLElement>('.docx-pages')!;
  // happy-dom lays nothing out, so the one rectangle the pointer reads is stated.
  Object.defineProperty(pages, 'getBoundingClientRect', {
    configurable: true,
    value: () => ({
      left: 0,
      top: 0,
      right: 2000,
      bottom: 4000,
      width: 2000,
      height: 4000,
      x: 0,
      y: 0,
    }),
  });
  const page = surface.layout().pages[0]!;
  let point: { clientX: number; clientY: number } | null = null;
  const scale = editor.getRenderScale();
  for (const fragment of page.fragments) {
    if (fragment.kind !== 'paragraph') continue;
    for (const line of fragment.lines) {
      const span = line.spans.find((candidate) => candidate.fieldAtom !== undefined);
      if (!span) continue;
      point = {
        clientX: (page.contentBox.x + span.box.x + span.box.width / 2) * scale,
        clientY: (page.contentBox.y + span.box.y + span.box.height / 2) * scale,
      };
    }
  }
  if (!point) throw new Error('no field painted');
  const init = { bubbles: true, cancelable: true, button: 0, pointerId: 1, pointerType: 'mouse' };
  pages.dispatchEvent(new PointerEvent('pointerdown', { ...init, ...point }));
  document.dispatchEvent(new PointerEvent('pointerup', { ...init, ...point }));
}

function installView(editor: DocxEditorInstance): void {
  editor.setContentControlTags({
    labelsOf: ({ tag }) => ({ open: { text: `${tag} ▸` }, close: { text: '◂' } }),
  });
  editor.setFieldTones((instruction) => (instruction.includes('"field:') ? 'variable' : undefined));
  editor.setFieldSelection((instruction) => instruction.includes('"field:'));
}

describe('the host view on the editor', () => {
  test('a load paints the next document with the tags and the tones', () => {
    const { editor, host } = mounted();
    installView(editor);
    expect(chipsIn(host)).toEqual(['a ▸', '◂']);
    editor.load(SECOND);
    expect(chipsIn(host)).toEqual(['b ▸', '◂']);
    expect(tonesIn(host)).toContain('variable');
  });

  test('a detach and an attach keep the view', () => {
    const { editor, host } = mounted();
    installView(editor);
    editor.detach();
    editor.attach(host);
    expect(chipsIn(host)).toEqual(['a ▸', '◂']);
    expect(tonesIn(host)).toContain('variable');
  });

  test('the field selection reaches a surface the editor mounts after it was set', () => {
    const { editor, host } = mounted();
    installView(editor);
    editor.load(SECOND);
    pressField(editor, host);
    const { anchor, head } = editor.surface!.state().selection;
    expect([anchor.offset, head.offset]).toEqual([6, 7]);
  });

  test('null takes each part of the view away, on the live surface and on the next one', () => {
    const { editor, host } = mounted();
    installView(editor);
    editor.setContentControlTags(null);
    editor.setFieldTones(null);
    expect(chipsIn(host)).toEqual([]);
    editor.load(SECOND);
    expect(chipsIn(host)).toEqual([]);
    expect(tonesIn(host).every((tone) => tone === undefined)).toBe(true);
  });

  test('the prompts reach the next document and never the saved bytes', async () => {
    const { editor, host } = mounted();
    editor.setContentControlPrompts({ promptOf: ({ tag }) => (tag === 'p' ? 'digite' : null) });
    editor.load(PROMPTED);
    expect(paintedText(host)).toContain('tres digite');
    expect(paintedText(host)).not.toContain('Clique aqui');
    const saved = unzipSync(new Uint8Array(await editor.save()))['word/document.xml']!;
    expect(strFromU8(saved)).toContain('Clique aqui');
    expect(strFromU8(saved)).not.toContain('digite');
    editor.setContentControlPrompts(null);
    expect(paintedText(host)).toContain('tres Clique aqui');
  });
});
