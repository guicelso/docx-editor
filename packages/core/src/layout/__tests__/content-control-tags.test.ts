// View-only content-control tags: they paint and take width, and they never move an offset.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, type OoxmlPart } from '../../store/package/ooxml-tree.ts';
import { createFixedMeasurer, layoutSemanticDocument } from '../semantic-layout.ts';
import type { StyleSpanRecord } from '../semantic-records.ts';
import { forEachSemanticSpan } from '../export-traversal.ts';
import { hitTestPage } from '../semantic-hit-test.ts';
import { anchorLineStartsByModelOffset } from '../anchor-line-probe.ts';
import { DEFAULT_RUN_STYLE } from '../run-style.ts';
import { contentControlTagChromePt } from '../content-control-tags.ts';
import type { ContentControlTagDisplay } from '../../contracts/editor-content-control-view.ts';
import {
  caretAt,
  caretStops,
  contentControlAtSemantic,
  contentControlsInLayout,
  hitTestSemantic,
} from '../semantic-interaction.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function documentOf(bodyXml: string): OoxmlPart {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body>${bodyXml}</w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'app/xml' }
  );
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const sdt = (tag: string, inner: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

// "CPF " + group{ branch{"RG"} else{"CNH"} } + " fim" — the shape of a block body.
const BODY = `<w:p>${run('CPF ')}${sdt(
  'span:alternatives:g1',
  sdt('span:branch:b1', run('RG')) + sdt('span:else:e1', run('CNH'))
)}${run(' fim')}</w:p>`;

const LABELS: Record<string, readonly [string, string]> = {
  'span:alternatives:g1': ['[Grupo', 'Grupo]'],
  'span:branch:b1': ['[Ramo 1', ']'],
  'span:else:e1': ['[Senão', ']'],
};

const TAGS: ContentControlTagDisplay = {
  labelsOf: ({ tag }) => {
    const pair = tag === undefined ? undefined : LABELS[tag];
    return pair ? { open: { text: pair[0] }, close: { text: pair[1] } } : null;
  },
};

const measurer = createFixedMeasurer(6, 14);
const layoutOf = (tags?: ContentControlTagDisplay) =>
  layoutSemanticDocument(documentOf(BODY), 0, {
    measurer,
    ...(tags ? { contentControlView: { tags } } : {}),
  });

function spansOf(layout: ReturnType<typeof layoutSemanticDocument>): StyleSpanRecord[] {
  const spans: StyleSpanRecord[] = [];
  for (const page of layout.pages)
    for (const fragment of page.fragments)
      for (const line of (fragment as { lines?: { spans: StyleSpanRecord[] }[] }).lines ?? [])
        spans.push(...line.spans);
  return spans;
}

/** A tag's painted text with its spaces as typed. */
const plain = (text: string) => text.replace(/\u00A0/g, ' ');

const paragraphIdOf = (layout: ReturnType<typeof layoutSemanticDocument>) =>
  spansOf(layout)[0]!.range.paragraphId;

describe('content-control tags', () => {
  test('paint in reading order, nested, around the content they mark', () => {
    const painted = spansOf(layoutOf(TAGS))
      .map((span) => span.text)
      .join('');
    expect(plain(painted)).toBe('CPF [Grupo[Ramo 1RG][SenãoCNH]Grupo] fim');
  });

  test('without a display, nothing is drawn and the layout is the untagged one', () => {
    expect(
      spansOf(layoutOf())
        .map((span) => span.text)
        .join('')
    ).toBe('CPF RGCNH fim');
  });

  test('every text span keeps exactly the model range it had without tags', () => {
    const ranges = (layout: ReturnType<typeof layoutSemanticDocument>) =>
      spansOf(layout)
        .filter((span) => !span.contentControlTag)
        .map((span) => `${span.text}@${span.range.start}-${span.range.end}`);
    expect(ranges(layoutOf(TAGS))).toEqual(ranges(layoutOf()));
  });

  test('a tag covers a ZERO-WIDTH range at its control edge and takes real width', () => {
    const tags = spansOf(layoutOf(TAGS)).filter((span) => span.contentControlTag);
    expect(
      tags.map((span) => `${span.contentControlTag!.edge}:${span.range.start}-${span.range.end}`)
    ).toEqual([
      'open:4-4', // group opens at "CPF |"
      'open:4-4', // branch opens at the same offset
      'close:6-6', // branch closes after "RG"
      'open:6-6', // else opens there
      'close:9-9', // else closes after "CNH"
      'close:9-9', // group closes there
    ]);
    for (const span of tags) expect(span.box.width).toBeGreaterThan(0);
  });

  test('a label is one chip at every break opportunity: a dash, a tab, a line too narrow', () => {
    const label = (text: string): ContentControlTagDisplay => ({
      labelsOf: () => ({ open: { text }, close: { text } }),
    });
    for (const text of ['group-1 ▸', 'a\tb', `wide-${'x'.repeat(200)}`]) {
      const chips = spansOf(layoutOf(label(text))).filter((span) => span.contentControlTag);
      expect(chips.map((span) => plain(span.text).trim())).toHaveLength(6);
      for (const chip of chips) expect(plain(chip.text).trim()).toBe(text);
    }
  });

  test('a chip measures its label plus the room around it, and the label carries no spacing characters', () => {
    const chips = spansOf(layoutOf(TAGS)).filter((span) => span.contentControlTag);
    for (const chip of chips) {
      expect(chip.text).not.toMatch(/\u2009/);
      expect(chip.box.width).toBeCloseTo(
        measurer.measure(chip.text, chip.style) +
          contentControlTagChromePt(chip.style, chip.contentControlTag!),
        6
      );
    }
  });

  test('a chip has the face, size and band of the text inside its control, not of its neighbours', () => {
    const sized = (halfPoints: number, text: string, extra = '') =>
      `<w:r><w:rPr>${extra}<w:sz w:val="${halfPoints}"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
    const body = `<w:p>${sized(20, 'CPF ')}${sdt('span:alternatives:g1', sized(28, 'RG', '<w:b/>'))}${sized(16, ' fim')}</w:p>`;
    const spans = spansOf(
      layoutSemanticDocument(documentOf(body), 0, { measurer, contentControlView: { tags: TAGS } })
    );
    const text = spans.find((span) => span.text === 'RG')!;
    const chips = spans.filter((span) => span.contentControlTag);
    expect(chips).toHaveLength(2);
    for (const chip of chips) {
      expect(chip.style.fontSizePt).toBe(14);
      expect(chip.style.bold).toBe(false);
      expect(chip.box.height).toBeCloseTo(text.box.height, 6);
    }
  });

  test('a chip with no text inside its control takes the text on its other side', () => {
    const body = `<w:p><w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:t>CPF</w:t></w:r>${sdt('span:alternatives:g1', '')}</w:p>`;
    const chips = spansOf(
      layoutSemanticDocument(documentOf(body), 0, { measurer, contentControlView: { tags: TAGS } })
    ).filter((span) => span.contentControlTag);
    expect(chips.map((chip) => chip.style.fontSizePt)).toEqual([14, 14]);
  });

  test('tags push the following text right; the line is wider by their width', () => {
    const plain = spansOf(layoutOf()).find((span) => span.text === 'fim')!;
    const tagged = spansOf(layoutOf(TAGS)).find((span) => span.text === 'fim')!;
    const tagWidth = spansOf(layoutOf(TAGS))
      .filter((span) => span.contentControlTag)
      .reduce((sum, span) => sum + span.box.width, 0);
    expect(tagged.box.x - plain.box.x).toBeCloseTo(tagWidth, 3);
  });

  test('the caret inside the text sits AFTER the tags before it', () => {
    const tagged = layoutOf(TAGS);
    const id = paragraphIdOf(tagged);
    const rg = spansOf(tagged).find((span) => span.text === 'RG')!;
    // Offset 5 is between "R" and "G": strictly inside the branch, never ambiguous.
    const caret = caretAt(tagged, { paragraphId: id, offset: 5 }, measurer)!;
    expect(caret.x).toBeGreaterThan(rg.box.x);
    expect(caret.x).toBeLessThan(rg.box.x + rg.box.width);
  });

  test('a click on a tag lands on its edge offset; a click on text lands where it did', () => {
    const tagged = layoutOf(TAGS);
    const spans = spansOf(tagged);
    const elseTag = spans.find((span) => plain(span.text) === '[Senão')!;
    const hit = hitTestSemantic(tagged, {
      x: elseTag.box.x + elseTag.box.width / 2,
      y: elseTag.box.y + elseTag.box.height / 2,
    })!;
    expect(hit.position.offset).toBe(6);

    const cnh = spans.find((span) => span.text === 'CNH')!;
    const onN = hitTestSemantic(tagged, {
      x: cnh.box.x + measurer.measure('C', cnh.style) * 1.2,
      y: cnh.box.y + cnh.box.height / 2,
    })!;
    expect(onN.position.offset).toBe(7);
  });

  test('caret stops: every model offset still has a stop', () => {
    const tagged = layoutOf(TAGS);
    const offsets = caretStops(tagged, measurer).map((stop) => stop.position.offset);
    for (let offset = 0; offset <= 13; offset += 1) expect(offsets).toContain(offset);
  });

  test("each outline wraps its own tags and its children, never a sibling's", () => {
    const tagged = layoutOf(TAGS);
    const spans = spansOf(tagged);
    const box = (text: string) => spans.find((span) => plain(span.text) === text)!.box;
    const outline = (tag: string) => {
      const record = contentControlsInLayout(tagged).find((control) => control.tag === tag)!;
      expect(record.fragments).toHaveLength(1);
      return record.fragments[0]!.box;
    };
    const group = outline('span:alternatives:g1');
    const branch = outline('span:branch:b1');
    const otherwise = outline('span:else:e1');
    expect(group.x).toBeCloseTo(box('[Grupo').x, 3);
    expect(group.x + group.width).toBeCloseTo(box('Grupo]').x + box('Grupo]').width, 3);
    expect(branch.x).toBeCloseTo(box('[Ramo 1').x, 3);
    expect(branch.x + branch.width).toBeCloseTo(box('[Senão').x, 3);
    expect(otherwise.x).toBeCloseTo(box('[Senão').x, 3);
    expect(otherwise.x + otherwise.width).toBeCloseTo(box('Grupo]').x, 3);
  });

  test("a point on a tag is in that tag's control: the group tag is the group, not the branch", () => {
    const tagged = layoutOf(TAGS);
    const spans = spansOf(tagged);
    const at = (text: string) => {
      const b = spans.find((span) => plain(span.text) === text)!.box;
      return contentControlAtSemantic(tagged, { x: b.x + b.width / 2, y: b.y + b.height / 2 })?.tag;
    };
    expect(at('[Grupo')).toBe('span:alternatives:g1');
    expect(at('[Ramo 1')).toBe('span:branch:b1');
    expect(at('[Senão')).toBe('span:else:e1');
    expect(at('Grupo]')).toBe('span:alternatives:g1');
  });
});

describe('content-control tags in a table cell', () => {
  const cell = (inner: string) =>
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol/></w:tblGrid>` +
    `<w:tr><w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>${inner}</w:tc></w:tr></w:tbl><w:p/>`;
  const CELL = cell(`<w:p>${run('CPF ')}${sdt('span:optional:c1', run('RG'))}</w:p>`);

  const cellTexts = (tags?: ContentControlTagDisplay) => {
    const layout = layoutSemanticDocument(documentOf(CELL), 0, {
      measurer,
      ...(tags ? { contentControlView: { tags } } : {}),
    });
    const texts: string[] = [];
    forEachSemanticSpan(layout, ({ span }) => texts.push(span.text));
    return { layout, text: plain(texts.join('')) };
  };

  test('the cell draws its controls’ tags', () => {
    expect(cellTexts(TAGS_FOR_CELL).text).toBe('CPF [RG]');
  });

  test('a label stays one chip in an auto-width cell and beside CJK prose', () => {
    const hyphenated: ContentControlTagDisplay = {
      labelsOf: () => ({ open: { text: 'group-1' }, close: { text: 'group-1' } }),
    };
    for (const body of [CELL, `<w:p>${run('漢字 ')}${sdt('span:optional:c1', run('RG'))}</w:p>`]) {
      const layout = layoutSemanticDocument(documentOf(body), 0, {
        measurer,
        contentControlView: { tags: hyphenated },
      });
      const chips: string[] = [];
      forEachSemanticSpan(layout, ({ span }) => {
        if (span.contentControlTag) chips.push(plain(span.text).trim());
      });
      expect(chips).toEqual(['group-1', 'group-1']);
    }
  });

  test('a squeezed auto column keeps a whole chip as its minimum, never spilling into the next', () => {
    const autoCell = (inner: string) =>
      `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p>${inner}</w:p></w:tc>`;
    const table =
      `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>` +
      `<w:tr>${autoCell(sdt('c1', run('a')))}${autoCell(run('zzzz '.repeat(200)))}</w:tr></w:tbl><w:p/>`;
    const layout = layoutSemanticDocument(documentOf(table), 0, {
      measurer,
      contentControlView: {
        tags: {
          labelsOf: () => ({ open: { text: `wide-${'y'.repeat(12)}` }, close: { text: '◂' } }),
        },
      },
    });
    let chipRight = 0;
    let nextColumnLeft = Number.POSITIVE_INFINITY;
    forEachSemanticSpan(layout, ({ span, absoluteBox }) => {
      if (span.contentControlTag)
        chipRight = Math.max(chipRight, absoluteBox.x + absoluteBox.width);
      if (span.text.startsWith('zzzz')) nextColumnLeft = Math.min(nextColumnLeft, absoluteBox.x);
    });
    expect(chipRight).toBeLessThanOrEqual(nextColumnLeft);
  });

  test('an auto-width cell widens by its tags, so the line never overflows', () => {
    const widthOf = (tags?: ContentControlTagDisplay) => {
      let widest = 0;
      forEachSemanticSpan(cellTexts(tags).layout, ({ absoluteBox }) => {
        widest = Math.max(widest, absoluteBox.x + absoluteBox.width);
      });
      return widest;
    };
    expect(widthOf(TAGS_FOR_CELL)).toBeGreaterThan(widthOf());
    const tagged = cellTexts(TAGS_FOR_CELL).layout;
    const lines = new Set<string>();
    forEachSemanticSpan(tagged, ({ span }) => lines.add(String(span.box.y)));
    expect(lines.size).toBe(1);
  });
});

const TAGS_FOR_CELL: ContentControlTagDisplay = {
  labelsOf: () => ({ open: { text: '[' }, close: { text: ']' } }),
};

describe('a press on a tag reports the chip and its half', () => {
  test('the left half asks for the slot in front of the chip, the right half for the one behind', () => {
    const tagged = layoutOf(TAGS);
    const chip = spansOf(tagged).find((span) => plain(span.text) === '[Senão')!;
    const at = (x: number) =>
      hitTestPage(tagged, 0, { x, y: chip.box.y + chip.box.height / 2 })?.contentControlTag;
    expect(at(chip.box.x + chip.box.width * 0.25)).toEqual({
      controlId: chip.contentControlTag!.controlId,
      edge: 'open',
      side: 'before',
    });
    expect(at(chip.box.x + chip.box.width * 0.75)?.side).toBe('after');
  });

  test('a press on text reports no chip', () => {
    const tagged = layoutOf(TAGS);
    const cnh = spansOf(tagged).find((span) => span.text === 'CNH')!;
    expect(
      hitTestPage(tagged, 0, { x: cnh.box.x + 1, y: cnh.box.y + cnh.box.height / 2 })
        ?.contentControlTag
    ).toBeUndefined();
  });
});

describe('an anchored drawing beside a tag', () => {
  test('its line is predicted with the chip whole, as placement breaks it', () => {
    const text = (value: string, start: number) => ({
      text: value,
      start,
      end: start + value.length,
      props: [],
      style: DEFAULT_RUN_STYLE,
    });
    const chip = {
      ...text('bb-cc', 4),
      end: 4,
      projected: true,
      contentControlTag: { controlId: 'c1', edge: 'open' as const },
    };
    const starts = anchorLineStartsByModelOffset({
      pieces: [text('aaaa', 0), chip, text('dd', 4)],
      measurer: createFixedMeasurer(10, 12),
      available: 70,
      firstLineOffset: 0,
      anchorStarts: [4],
      equationLayoutOf: () => null,
    });
    expect(starts.get(4)).toBe(4);
  });
});
