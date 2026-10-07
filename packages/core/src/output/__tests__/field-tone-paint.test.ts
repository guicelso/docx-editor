// A host names its own fields for its stylesheet: the paint asks it, by instruction, and
// publishes the answer on the field's result.

import { GlobalRegistrator } from '@happy-dom/global-registrator';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, type OoxmlPart } from '@docx-editor.dev/core/store';
import { createFixedMeasurer, layoutSemanticDocument } from '../../layout/semantic-layout.ts';
import { paintSemanticLayout, type FieldTone } from '../semantic-paint.ts';

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

const TONES: FieldTone = (instruction) =>
  instruction.includes('"field:')
    ? 'variable'
    : instruction.includes('"block:')
      ? 'citation'
      : instruction.includes('"mark"')
        ? 'mark'
        : undefined;

function tones(fieldTone?: FieldTone): (string | undefined)[] {
  const layout = layoutSemanticDocument(load(BODY), 1, { measurer });
  const container = document.createElement('div');
  paintSemanticLayout(container, layout, {
    scale: 1,
    ariaHidden: false,
    ...(fieldTone ? { fieldTone } : {}),
  });
  return [...container.querySelectorAll<HTMLElement>('[data-field-atom]')].map(
    (field) => field.dataset.fieldTone
  );
}

describe('field tones', () => {
  test('each field’s result carries the name the host gave its instruction', () => {
    expect(tones(TONES)).toEqual(['variable', 'citation', 'mark']);
  });

  test('without a host, or with no answer, nothing is published', () => {
    expect(tones()).toEqual([undefined, undefined, undefined]);
    expect(tones(() => undefined)).toEqual([undefined, undefined, undefined]);
  });

  test('a name that is not a plain token never reaches the attribute', () => {
    expect(tones(() => 'x" onclick="1')).toEqual([undefined, undefined, undefined]);
  });
});
