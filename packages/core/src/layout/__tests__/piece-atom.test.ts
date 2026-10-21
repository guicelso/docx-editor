// An atom is a piece the reader meets as one unit: a field with the host's face, or a content
// control the host made an atom. It measures its text plus a chip's room and its mark, never
// breaks, and keeps every offset of the store.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, type OoxmlPart } from '../../store/package/ooxml-tree.ts';
import type {
  ContentControlPieceDisplay,
  FieldFace,
} from '../../contracts/editor-content-control-view.ts';
import { createFixedMeasurer, layoutSemanticDocument } from '../semantic-layout.ts';
import type { SemanticLayout, StyleSpanRecord } from '../semantic-records.ts';
import type { ContentControlView } from '../content-control-view.ts';
import { pieceRoomOf, pieceRoomWidthPt } from '../piece-names.ts';
import { caretAt, contentControlsInLayout } from '../semantic-interaction.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function documentOf(bodyXml: string): OoxmlPart {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body>${bodyXml}</w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'app/xml' }
  );
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

const run = (text: string, properties = '') =>
  `<w:r>${properties === '' ? '' : `<w:rPr>${properties}</w:rPr>`}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const field = (instruction: string, result: string) =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
  `<w:r><w:instrText xml:space="preserve">${instruction}</w:instrText></w:r>` +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${run(result)}` +
  `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
const sdt = (tag: string, inner: string, properties = '') =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/>${properties}</w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

const measurer = createFixedMeasurer(6, 14);
const layoutOf = (body: string, view?: ContentControlView): SemanticLayout =>
  layoutSemanticDocument(documentOf(body), 0, {
    measurer,
    ...(view ? { contentControlView: view } : {}),
  });

function spansOf(layout: SemanticLayout): StyleSpanRecord[] {
  const spans: StyleSpanRecord[] = [];
  for (const page of layout.pages)
    for (const fragment of page.fragments)
      for (const line of (fragment as { lines?: { spans: StyleSpanRecord[] }[] }).lines ?? [])
        spans.push(...line.spans);
  return spans;
}

const textOf = (layout: SemanticLayout): string =>
  spansOf(layout)
    .map((span) => span.text)
    .join('');

const VARIABLE: FieldFace = (instruction) =>
  instruction.includes('MERGEFIELD')
    ? { tone: 'variable', form: 'atom', mark: 'variable' }
    : undefined;

const FIELD_BODY = `<w:p>${run('O ')}${field(' MERGEFIELD "v:1" ', '«Nome do comprador»')}${run(' fim')}</w:p>`;

describe('a field with a face', () => {
  test('is ONE atom: its result, a chip’s room on each side and its mark', () => {
    const atoms = spansOf(layoutOf(FIELD_BODY, { fieldFaces: VARIABLE })).filter(
      (span) => span.fieldAtom
    );
    expect(atoms).toHaveLength(1);
    const atom = atoms[0]!;
    expect(atom.pieceAtom).toEqual({ name: { tone: 'variable', form: 'atom' }, mark: 'variable' });
    const room = pieceRoomOf(atom.style, 'variable');
    expect(room.markPt).toBeCloseTo(atom.style.fontSizePt * 1.15, 6);
    expect(atom.box.width).toBeCloseTo(
      measurer.measure(atom.text, atom.style) + pieceRoomWidthPt(room),
      6
    );
  });

  test('without a face the field is what it was: its result breaks at its spaces', () => {
    const results = spansOf(layoutOf(FIELD_BODY)).filter((span) => span.fieldAtom);
    expect(results.length).toBeGreaterThan(1);
    for (const span of results) expect(span.pieceAtom).toBeUndefined();
  });

  test('never breaks: on a line too narrow for it, the whole atom moves, mark and result together', () => {
    const narrow = `<w:p>${run('x '.repeat(30))}${field(' MERGEFIELD "v:1" ', '«Nome do comprador»')}</w:p>`;
    const layout = layoutSemanticDocument(documentOf(narrow), 0, {
      measurer,
      contentControlView: { fieldFaces: VARIABLE },
      geometry: { width: 300, height: 800, margin: { top: 20, right: 20, bottom: 20, left: 20 } },
    });
    const atoms = spansOf(layout).filter((span) => span.fieldAtom);
    expect(atoms).toHaveLength(1);
    expect(atoms[0]!.text).toBe('«Nome do comprador»');
  });

  test('the caret stands at its painted edges: before the mark and after the room', () => {
    const layout = layoutOf(FIELD_BODY, { fieldFaces: VARIABLE });
    const atom = spansOf(layout).find((span) => span.fieldAtom)!;
    const { paragraphId, start, end } = atom.range;
    expect(caretAt(layout, { paragraphId, offset: start }, measurer)!.x).toBeCloseTo(atom.box.x, 6);
    expect(caretAt(layout, { paragraphId, offset: end }, measurer)!.x).toBeCloseTo(
      atom.box.x + atom.box.width,
      6
    );
  });

  test('a name or a mark outside the pattern is neither published nor measured', () => {
    const hostile: FieldFace = () => ({ tone: 'x" onclick="1', form: 'atom', mark: '<svg>' });
    const atom = spansOf(layoutOf(FIELD_BODY, { fieldFaces: hostile })).find(
      (span) => span.fieldAtom
    )!;
    expect(atom.pieceAtom).toEqual({ name: { form: 'atom' } });
    expect(atom.box.width).toBeCloseTo(
      measurer.measure(atom.text, atom.style) +
        pieceRoomWidthPt(pieceRoomOf(atom.style, undefined)),
      6
    );
  });
});

const ATOM: ContentControlPieceDisplay = {
  pieceOf: ({ tag }) =>
    tag === 'entry:q1'
      ? { atom: { tone: 'citation', form: 'atom', mark: 'qualification' } }
      : tag === 'entry:b1'
        ? { frame: { tone: 'citation', form: 'embed' } }
        : null,
};

const QUALIFICATION = `<w:p>${run('A ')}${sdt('entry:q1', run('Qualificação completa · Vendedor(a)'))}${run(' fim')}</w:p>`;

describe('a content control the host made an atom', () => {
  test('lays its plain text out as ONE projected piece over the content’s range', () => {
    const layout = layoutOf(QUALIFICATION, { pieces: ATOM });
    const control = contentControlsInLayout(layout).find((each) => each.tag === 'entry:q1')!;
    const atoms = spansOf(layout).filter((span) => span.pieceAtom);
    expect(atoms).toHaveLength(1);
    const atom = atoms[0]!;
    expect(atom.text).toBe('Qualificação completa · Vendedor(a)');
    expect(atom.projected).toBe(true);
    expect(atom.range.end - atom.range.start).toBe(atom.text.length);
    expect(atom.pieceAtom).toEqual({
      name: { tone: 'citation', form: 'atom' },
      mark: 'qualification',
      controlId: control.id,
    });
    expect(atom.box.width).toBeCloseTo(
      measurer.measure(atom.text, atom.style) +
        pieceRoomWidthPt(pieceRoomOf(atom.style, 'qualification')),
      6
    );
  });

  test('keeps every offset the store has: the text around it lands where it did', () => {
    const ranges = (layout: SemanticLayout) =>
      spansOf(layout)
        .filter((span) => span.text === 'A ' || span.text === ' fim')
        .map((span) => `${span.text}@${span.range.start}-${span.range.end}`);
    expect(ranges(layoutOf(QUALIFICATION, { pieces: ATOM }))).toEqual(
      ranges(layoutOf(QUALIFICATION))
    );
  });

  test('its record says it is an atom, with the host’s names; a frame only names its chrome', () => {
    const body = `${QUALIFICATION}<w:p>${sdt('entry:b1', run('Texto do bloco'))}</w:p>`;
    const records = contentControlsInLayout(layoutOf(body, { pieces: ATOM }));
    const atom = records.find((each) => each.tag === 'entry:q1')!;
    const frame = records.find((each) => each.tag === 'entry:b1')!;
    expect(atom.atom).toBe(true);
    expect(atom.name).toEqual({ tone: 'citation', form: 'atom' });
    expect(frame.atom).toBe(false);
    expect(frame.name).toEqual({ tone: 'citation', form: 'embed' });
    expect(textOf(layoutOf(body, { pieces: ATOM }))).toContain('Texto do bloco');
  });

  test('text in more than one face keeps its pieces, and the control is still met as a unit', () => {
    const body = `<w:p>${sdt('entry:q1', run('Qualificação', '<w:b/>') + run(' completa'))}</w:p>`;
    const layout = layoutOf(body, { pieces: ATOM });
    expect(spansOf(layout).some((span) => span.pieceAtom)).toBe(false);
    expect(contentControlsInLayout(layout)[0]!.atom).toBe(true);
  });

  test('content the file marks — a field inside it — keeps its pieces', () => {
    const body = `<w:p>${sdt('entry:q1', run('Nome ') + field(' MERGEFIELD "v:1" ', '«x»'))}</w:p>`;
    expect(spansOf(layoutOf(body, { pieces: ATOM })).some((span) => span.pieceAtom)).toBe(false);
  });

  test('a control showing its placeholder shows its prompt, and is no atom', () => {
    const body = `<w:p>${sdt('entry:q1', run('Clique aqui'), '<w:showingPlcHdr/>')}</w:p>`;
    const layout = layoutOf(body, {
      pieces: ATOM,
      prompts: { promptOf: () => 'Escolha a qualificação' },
    });
    expect(spansOf(layout).some((span) => span.pieceAtom)).toBe(false);
    expect(textOf(layout)).toBe('Escolha a qualificação');
    expect(contentControlsInLayout(layout)[0]!.atom).toBe(false);
  });

  test('a block control holding one paragraph of plain text is an atom too', () => {
    const body = sdt('entry:q1', `<w:p>${run('Qualificação completa')}</w:p>`);
    const atoms = spansOf(layoutOf(body, { pieces: ATOM })).filter((span) => span.pieceAtom);
    expect(atoms.map((span) => span.text)).toEqual(['Qualificação completa']);
  });
});
