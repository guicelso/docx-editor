// A view-only tag is chrome, not a letter of the word beside it: a line may open beside it, and
// a tag no line can hold is cut to fit. Nothing a tag draws stands past the margin.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart } from '@docx-editor.dev/core/store';
import type { ContentControlTagDisplay } from '../../contracts/editor-content-control-view.ts';
import { opensAtTagSeam, withContentControlTagsFitted } from '../content-control-tag-lines.ts';
import type { ContentControlTagMark } from '../content-control-tags.ts';
import type { FieldAwarePiece } from '../field-pieces.ts';
import { DEFAULT_RUN_STYLE } from '../run-style.ts';
import { createFixedMeasurer, layoutSemanticDocument } from '../semantic-layout.ts';
import { linesOf } from '../semantic-records.ts';
import { elevenPointDefaults } from './fixtures/eleven-point-defaults.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STYLE = { ...DEFAULT_RUN_STYLE, fontSizePt: 11 };
const measurer = createFixedMeasurer(6, 12);

function textPiece(text: string): FieldAwarePiece {
  return { text, props: [], style: STYLE, start: 0, end: 0 };
}

function tagPiece(
  text: string,
  edge: ContentControlTagMark['edge'],
  mark: Partial<ContentControlTagMark> = {}
): FieldAwarePiece {
  return {
    ...textPiece(text),
    projected: true,
    contentControlTag: { controlId: 'c1', edge, level: 'inline', variant: 'text', ...mark },
  };
}

describe('opensAtTagSeam', () => {
  test('a line may open before an open tag and after a close tag', () => {
    expect(opensAtTagSeam(textPiece('palavra'), tagPiece('▸ Trecho', 'open'))).toBe(true);
    expect(opensAtTagSeam(tagPiece('◂', 'close'), textPiece('palavra'))).toBe(true);
    expect(opensAtTagSeam(tagPiece('◂', 'close'), tagPiece('▸ Trecho', 'open'))).toBe(true);
  });

  test('the tags of nested controls travel together', () => {
    expect(opensAtTagSeam(tagPiece('▸ Fora', 'open'), tagPiece('▸ Dentro', 'open'))).toBe(false);
    expect(opensAtTagSeam(tagPiece('◂', 'close'), tagPiece('◂', 'close'))).toBe(false);
  });

  test('punctuation that may not end or start a line stays with its tag', () => {
    expect(opensAtTagSeam(textPiece('('), tagPiece('▸ Trecho', 'open'))).toBe(false);
    expect(opensAtTagSeam(textPiece('«'), tagPiece('▸ Trecho', 'open'))).toBe(false);
    expect(opensAtTagSeam(tagPiece('◂', 'close'), textPiece('.'))).toBe(false);
    expect(opensAtTagSeam(tagPiece('◂', 'close'), textPiece('»'))).toBe(false);
    expect(opensAtTagSeam(tagPiece('◂', 'close'), textPiece(', e'))).toBe(false);
  });

  test('a seam between two pieces of text is not a tag seam', () => {
    expect(opensAtTagSeam(textPiece('a'), textPiece('b'))).toBe(false);
    expect(opensAtTagSeam(undefined, tagPiece('▸ Trecho', 'open'))).toBe(false);
  });
});

describe('withContentControlTagsFitted', () => {
  test('pieces without a tag come back as they were', () => {
    const pieces = [textPiece('a'), textPiece('b')];

    expect(withContentControlTagsFitted(pieces, 60, measurer)).toBe(pieces);
  });

  test('a tag that fits is not touched', () => {
    const tag = tagPiece('Trecho', 'open');

    expect(withContentControlTagsFitted([tag], 60, measurer)[0]).toBe(tag);
  });

  test('a tag wider than the measure keeps its start and ends in an ellipsis', () => {
    const [fitted] = withContentControlTagsFitted(
      [tagPiece('Forma de aquisição de cada Objeto', 'open')],
      60,
      measurer
    );

    expect(fitted!.text).toBe('Forma de…');
    expect(measurer.measure(fitted!.text, STYLE)).toBeLessThanOrEqual(60);
  });

  test('an open tag keeps room for the first word it wraps', () => {
    const [fitted] = withContentControlTagsFitted(
      [tagPiece('Forma de aquisição de cada Objeto', 'open'), textPiece('texto seguinte')],
      60,
      measurer
    );

    expect(fitted!.text).toBe('Form…');
  });

  test('two wide tags that travel together share the line, and neither gives up all its room', () => {
    const fitted = withContentControlTagsFitted(
      [tagPiece('a'.repeat(20), 'open'), tagPiece('b'.repeat(20), 'open'), textPiece('x')],
      60,
      measurer
    );

    expect(fitted.map((piece) => piece.text)).toEqual(['aaa…', 'bbb…', 'x']);
  });

  test('in a sequence only the longest label is cut, and the parts still cover the text', () => {
    const [fitted] = withContentControlTagsFitted(
      [
        tagPiece('▸ Ab cdefghijklmnopqrstuvwxyz', 'open', {
          variant: 'sequence',
          parts: [
            { length: 2, variant: 'text' },
            { length: 27, variant: 'text' },
          ],
        }),
      ],
      60,
      measurer
    );
    const parts = fitted!.contentControlTag!.parts!;

    expect(fitted!.text).toBe('▸ Ab cdef…');
    expect(parts.map((part) => part.length)).toEqual([2, 8]);
    expect(parts.reduce((sum, part) => sum + part.length, 0)).toBe(fitted!.text.length);
  });
});

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const sdt = (tag: string, inner: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

function layoutWithTags(content: string, labels: Record<string, readonly [string, string]>) {
  const tags: ContentControlTagDisplay = {
    labelsOf: ({ tag }) => {
      const pair = tag === undefined ? undefined : labels[tag];
      return pair
        ? { open: { text: pair[0], variant: 'text' }, close: { text: pair[1], variant: 'text' } }
        : null;
    },
  };
  const read = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body><w:p>${content}</w:p></w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'app/xml' }
  );
  if (!read.ok) throw new Error(read.reason);
  return layoutSemanticDocument(read.part, 0, {
    measurer,
    styleCascade: elevenPointDefaults(),
    contentControlView: { tags },
    geometry: { width: 120, height: 400, margin: { top: 0, bottom: 0, left: 0, right: 0 } },
  });
}

const rightEdgesOf = (result: ReturnType<typeof layoutWithTags>) =>
  linesOf(result).flatMap((line) => line.spans.map((span) => span.box.x + span.box.width));

describe('tags in a line', () => {
  test('a chain of tags with no space between them breaks at the seams, inside the margin', () => {
    const result = layoutWithTags(sdt('a', run('a')) + sdt('b', run('b')) + sdt('c', run('c')), {
      a: ['▸ Primeiro', '◂'],
      b: ['▸ Segundo', '◂'],
      c: ['▸ Terceiro', '◂'],
    });
    const lines = linesOf(result).map((line) => line.spans.map((span) => span.text).join(''));

    expect(lines.length).toBeGreaterThan(1);
    expect(lines.slice(1).every((line) => line.startsWith('▸'))).toBe(true);
    for (const edge of rightEdgesOf(result)) expect(edge).toBeLessThanOrEqual(120 + 0.01);
  });

  test('a tag wider than the line is cut to fit it beside the word it wraps', () => {
    const result = layoutWithTags(sdt('a', run('texto')), {
      a: ['▸ Forma de aquisição · de cada Objeto do imóvel', '◂'],
    });
    const opening = linesOf(result)
      .flatMap((line) => line.spans)
      .find((span) => span.contentControlTag?.edge === 'open');

    expect(opening?.text.endsWith('…')).toBe(true);
    for (const edge of rightEdgesOf(result)) expect(edge).toBeLessThanOrEqual(120 + 0.01);
  });
});
