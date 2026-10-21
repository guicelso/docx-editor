// A host gives its own fields a face: layout asks it, by instruction, and the field becomes an atom
// that paints the host's names and a measured slot for its mark.

import { GlobalRegistrator } from '@happy-dom/global-registrator';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, type OoxmlPart } from '@docx-editor.dev/core/store';
import { createFixedMeasurer, layoutSemanticDocument } from '../../layout/semantic-layout.ts';
import type { FieldFace } from '../../contracts/editor-content-control-view.ts';
import type { MarkPainter } from '../semantic-paint-pieces.ts';
import { PIECE_MARK_GLYPH_EM, PIECE_MARK_SPACE_EM } from '../../layout/piece-names.ts';
import { paintSemanticLayout } from '../semantic-paint.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const measurer = createFixedMeasurer(6, 14);

function load(body: string): OoxmlPart {
  const result = readOoxmlPart(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`, {
    name: '/word/document.xml',
    contentType: 'app/xml',
  });
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

const complex = (instruction: string, result: string) =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> ${instruction} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  `<w:r><w:t>${result}</w:t></w:r>` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

const BODY =
  `<w:p>${complex('MERGEFIELD "field:a"', '«cpf»')}<w:r><w:t xml:space="preserve"> e </w:t></w:r>` +
  `${complex('MERGEFIELD "block:b@SELF"', '⟦bloco⟧')}` +
  `<w:fldSimple w:instr=" MERGEFIELD &quot;mark&quot; "><w:r><w:t>◆</w:t></w:r></w:fldSimple></w:p>`;

const FACES: FieldFace = (instruction) =>
  instruction.includes('"field:')
    ? { tone: 'variable', form: 'atom', mark: 'variable' }
    : instruction.includes('"block:')
      ? { tone: 'citation' }
      : instruction.includes('"mark"')
        ? { tone: 'mark' }
        : undefined;

function painted(fieldFaces?: FieldFace, markPainter?: MarkPainter): HTMLElement[] {
  const layout = layoutSemanticDocument(load(BODY), 1, {
    measurer,
    ...(fieldFaces ? { contentControlView: { fieldFaces } } : {}),
  });
  const container = document.createElement('div');
  paintSemanticLayout(container, layout, {
    scale: 2,
    ariaHidden: false,
    ...(markPainter ? { markPainter } : {}),
  });
  return [...container.querySelectorAll<HTMLElement>('[data-field-atom]')];
}

describe('field faces', () => {
  test('each field’s result carries the names the host gave its instruction', () => {
    expect(painted(FACES).map((field) => field.dataset.pieceTone)).toEqual([
      'variable',
      'citation',
      'mark',
    ]);
    expect(painted(FACES)[0]!.dataset.pieceForm).toBe('atom');
  });

  test('without a host, or with no answer, nothing is published and nothing is measured', () => {
    for (const fields of [painted(), painted(() => undefined)]) {
      for (const field of fields) {
        expect(field.dataset.pieceTone).toBeUndefined();
        expect(field.querySelector('[data-piece-mark]')).toBeNull();
        expect(field.style.paddingLeft).toBe('');
      }
    }
  });

  test('a name that is not a plain token never reaches the attribute', () => {
    for (const field of painted(() => ({ tone: 'x" onclick="1', mark: '<b>' }))) {
      expect(field.dataset.pieceTone).toBeUndefined();
      expect(field.querySelector('[data-piece-mark]')).toBeNull();
    }
  });

  test('the mark slot is the size layout measured, before the result, and the painter fills it', () => {
    const [field] = painted(FACES, (mark, slot) => {
      slot.dataset.painted = mark;
    });
    const slot = field!.firstElementChild as HTMLElement;
    const em = parseFloat(field!.style.fontSize);
    expect(slot.classList.contains('docx-piece-mark')).toBe(true);
    expect(slot.dataset.pieceMark).toBe('variable');
    expect(slot.dataset.painted).toBe('variable');
    expect(slot.getAttribute('aria-hidden')).toBe('true');
    expect(parseFloat(slot.style.width)).toBeCloseTo(PIECE_MARK_GLYPH_EM * em, 4);
    expect(parseFloat(slot.style.marginRight)).toBeCloseTo(PIECE_MARK_SPACE_EM * em, 4);
    expect(field!.textContent).toBe('«cpf»');
  });

  test('without a painter the slot stays empty', () => {
    const [field] = painted(FACES);
    expect(field!.querySelector('[data-piece-mark]')?.childElementCount).toBe(0);
  });
});
