// A field is asked by its subject, as a control is: its start node names the occurrence, so a host
// answers two fields with the same instruction apart, and edits around a field keep its name.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import type { FieldFace } from '../../contracts/editor-content-control-view.ts';
import { paragraphRangeRects } from '../../layout/paragraph-range-rects.ts';
import { createFixedMeasurer, layoutSemanticDocument } from '../../layout/semantic-layout.ts';
import type { SemanticLayout, StyleSpanRecord } from '../../layout/semantic-records.ts';
import { parsedFieldSpansOf } from '../../store/package/field-nodes.ts';
import {
  readOoxmlPart,
  type OoxmlParagraphNode,
  type OoxmlPart,
} from '../../store/package/ooxml-tree.ts';
import type { PaginatedSurface } from '../paginated-surface.ts';
import { mount } from './paginated-surface-fixtures.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const text = (value: string) => `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`;
const field = (name: string) =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> MERGEFIELD "${name}" </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  text('«x»') +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
const simple = (name: string) =>
  `<w:fldSimple w:instr=" MERGEFIELD &quot;${name}&quot; ">${text('«y»')}</w:fldSimple>`;

const TWINS = `<w:p>${text('A ')}${field('a')}${text(' e ')}${field('a')}${text(' fim')}</w:p>`;

const mounted: PaginatedSurface[] = [];
afterEach(() => {
  for (const surface of mounted.splice(0)) surface.destroy();
});

function firstParagraphOf(part: OoxmlPart): OoxmlParagraphNode {
  const paragraph = part.root.children
    .flatMap((child) => (child.kind === 'textValue' ? [] : child.children))
    .find((node): node is OoxmlParagraphNode => node.kind === 'paragraph');
  if (!paragraph) throw new Error('no paragraph');
  return paragraph;
}

const fieldIdsOf = (part: OoxmlPart): string[] =>
  parsedFieldSpansOf(firstParagraphOf(part)).map((span) => span.node.id);

function documentOf(bodyXml: string): OoxmlPart {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body>${bodyXml}</w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'app/xml' }
  );
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

const measurer = createFixedMeasurer(6, 14);

function atomsOf(layout: SemanticLayout): StyleSpanRecord[] {
  const spans: StyleSpanRecord[] = [];
  for (const page of layout.pages)
    for (const fragment of page.fragments)
      for (const line of (fragment as { lines?: { spans: StyleSpanRecord[] }[] }).lines ?? [])
        spans.push(...line.spans);
  return spans.filter((span) => span.fieldAtom !== undefined);
}

describe('a field’s identity', () => {
  test('survives typing before, between and right against the fields of its paragraph', () => {
    const opened = mount(`<w:p>${text('A ')}${field('a')}${field('b')}${text(' fim')}</w:p>`);
    mounted.push(opened.surface);
    const before = fieldIdsOf(opened.surface.session.part());
    const paragraphId = opened.surface.state().selection.head.paragraphId;

    for (const offset of [0, 2, 3, 4]) {
      opened.surface.setSelection({
        anchor: { paragraphId, offset },
        head: { paragraphId, offset },
      });
      opened.surface.type('z');
      expect(fieldIdsOf(opened.surface.session.part())).toEqual(before);
    }
  });

  test('names the begin of a complex field and the element of a simple one', () => {
    const part = documentOf(`<w:p>${field('a')}${simple('b')}</w:p>`);
    const ids = fieldIdsOf(part);
    const layout = layoutSemanticDocument(part, 0, { measurer });

    const painted = atomsOf(layout).map((span) => span.fieldAtom?.field?.fieldId);
    expect(painted.toSorted()).toEqual(ids.toSorted());
  });
});

describe('a face asked by the subject', () => {
  test('tells two fields with the same instruction apart', () => {
    const part = documentOf(TWINS);
    const [, second] = fieldIdsOf(part);
    const face: FieldFace = ({ fieldId }) =>
      fieldId === second ? { tone: 'variable', state: 'selected' } : { tone: 'variable' };
    const atoms = atomsOf(
      layoutSemanticDocument(part, 0, { measurer, contentControlView: { fieldFaces: face } })
    );

    expect(atoms.map((atom) => atom.pieceAtom?.name.state)).toEqual([undefined, 'selected']);
  });
});

describe('a highlight of a field', () => {
  test('covers the field whole, from edge to edge, even with a face', () => {
    const face: FieldFace = () => ({ tone: 'variable', form: 'atom', mark: 'variable' });
    const layout = layoutSemanticDocument(documentOf(TWINS), 0, {
      measurer,
      contentControlView: { fieldFaces: face },
    });
    const [atom] = atomsOf(layout);
    const { paragraphId, start, end } = atom!.range;

    const rects = paragraphRangeRects(
      layout,
      new Map([[paragraphId, [{ key: 0, start, end }]]]),
      undefined,
      measurer
    );

    expect(rects).toHaveLength(1);
    expect(rects[0]!.x).toBeCloseTo(atom!.box.x, 6);
    expect(rects[0]!.width).toBeCloseTo(atom!.box.width, 6);
  });
});
