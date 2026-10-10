// The host's prompt in a content control showing its placeholder: it paints over the stored
// placeholder, in its style, as one unit, and the model offsets and text stay the stored ones.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, type OoxmlPart } from '../../store/package/ooxml-tree.ts';
import type { ContentControlPromptDisplay } from '../../contracts/editor-content-control-view.ts';
import { createParagraphLayoutCache } from '../layout-cache.ts';
import { createLayoutSession } from '../layout-session.ts';
import type { PendingLine } from '../paragraph-flow.ts';
import { createFixedMeasurer, layoutSemanticDocument } from '../semantic-layout.ts';
import type { StyleSpanRecord } from '../semantic-records.ts';
import { caretStops, paragraphTextFromLayout } from '../semantic-interaction.ts';
import type { ContentControlView } from '../content-control-view.ts';

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
const grey = (text: string) =>
  `<w:r><w:rPr><w:color w:val="808080"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
const control = (tag: string, inner: string, placeholder = true) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/>${placeholder ? '<w:showingPlcHdr/>' : ''}</w:sdtPr>` +
  `<w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
const paragraph = (inner: string) => `<w:p>${inner}</w:p>`;

const INLINE = paragraph(run('Nome: ') + control('a', grey('Clique aqui')) + run(' fim'));

const prompts = (answer: string | null): ContentControlPromptDisplay => ({
  promptOf: () => answer,
});

const measurer = createFixedMeasurer(6, 14);

function layoutOf(bodyXml: string, view?: ContentControlView) {
  return layoutSemanticDocument(documentOf(bodyXml), 0, {
    measurer,
    ...(view ? { contentControlView: view } : {}),
  });
}

function spansOf(layout: ReturnType<typeof layoutSemanticDocument>): StyleSpanRecord[] {
  const spans: StyleSpanRecord[] = [];
  for (const page of layout.pages)
    for (const fragment of page.fragments)
      for (const line of (fragment as { lines?: { spans: StyleSpanRecord[] }[] }).lines ?? [])
        spans.push(...line.spans);
  return spans;
}

const textOf = (layout: ReturnType<typeof layoutSemanticDocument>) =>
  spansOf(layout)
    .map((span) => span.text)
    .join('')
    .replace(/ /g, ' ');

describe('the prompt of a control showing its placeholder', () => {
  test('paints over the stored placeholder, in its style, over its whole range', () => {
    const stored = spansOf(layoutOf(INLINE)).find((span) => span.range.start === 6)!;
    expect(stored.style.color).not.toBeNull();
    const shown = layoutOf(INLINE, { prompts: prompts('digite o trecho') });
    expect(textOf(shown)).toBe('Nome: digite o trecho fim');
    // Cut at its words like a field result, every span publishing the whole stored range.
    const prompt = spansOf(shown).filter((span) => span.range.start === 6);
    expect(prompt.map((span) => span.text).join('')).toBe('digite o trecho');
    for (const span of prompt) {
      expect([span.range.end, span.projected, span.style.color]).toEqual([
        17,
        true,
        stored.style.color,
      ]);
    }
  });

  test('reads back as the stored text, so every offset agrees with the model', () => {
    const shown = layoutOf(INLINE, { prompts: prompts('digite o trecho') });
    const paragraphId = spansOf(shown)[0]!.range.paragraphId;
    expect(paragraphTextFromLayout(shown, paragraphId)).toBe('Nome: Clique aqui fim');
  });

  test('is one unit: the caret stops before and after it, never inside', () => {
    const shown = layoutOf(INLINE, { prompts: prompts('digite o trecho') });
    const offsets = caretStops(shown, measurer).map((stop) => stop.position.offset);
    expect(offsets).toContain(6);
    expect(offsets).toContain(17);
    expect(offsets.filter((offset) => offset > 6 && offset < 17)).toEqual([]);
  });

  test('null or an empty answer shows the stored placeholder', () => {
    expect(textOf(layoutOf(INLINE, { prompts: prompts(null) }))).toBe('Nome: Clique aqui fim');
    expect(textOf(layoutOf(INLINE, { prompts: prompts('') }))).toBe('Nome: Clique aqui fim');
  });

  test('a control holding its value shows the value', () => {
    const filled = paragraph(control('a', run('Maria'), false));
    expect(textOf(layoutOf(filled, { prompts: prompts('digite o trecho') }))).toBe('Maria');
  });

  test('a tracked placeholder keeps the stored text the revision marks', () => {
    const tracked = paragraph(
      control('a', `<w:ins w:id="1" w:author="Ana">${grey('Clique aqui')}</w:ins>`)
    );
    expect(textOf(layoutOf(tracked, { prompts: prompts('digite o trecho') }))).toBe('Clique aqui');
  });

  test('a block control shows it over its only paragraph, and its stored text over several', () => {
    const view = { prompts: prompts('digite o trecho') };
    const one = control('b', paragraph(grey('Clique aqui')));
    expect(textOf(layoutOf(one, view))).toBe('digite o trecho');
    const two = control('b', paragraph(grey('Clique')) + paragraph(grey('aqui')));
    expect(textOf(layoutOf(two, view))).toBe('Cliqueaqui');
  });

  test('draws no tag without tags, and sits between the tags with them', () => {
    const alone = spansOf(layoutOf(INLINE, { prompts: prompts('digite o trecho') }));
    expect(alone.some((span) => span.contentControlTag)).toBe(false);
    const tagged = layoutOf(INLINE, {
      prompts: prompts('digite o trecho'),
      tags: { labelsOf: () => ({ open: { text: '[' }, close: { text: ']' } }) },
    });
    expect(textOf(tagged)).toBe('Nome: [digite o trecho] fim');
  });

  test('a changed answer breaks only the paragraph that shows it', () => {
    const body =
      paragraph(control('a', grey('Clique aqui'))) +
      paragraph(control('b', grey('Clique aqui'))) +
      paragraph(run('tres'));
    const part = documentOf(body);
    const cache = createParagraphLayoutCache<readonly PendingLine[]>();
    const session = createLayoutSession();
    const layout = (answers: Readonly<Record<string, string>>) =>
      layoutSemanticDocument(part, 0, {
        measurer,
        cache,
        session,
        contentControlView: { prompts: { promptOf: ({ tag }) => answers[tag ?? ''] ?? null } },
      });
    layout({ a: 'um', b: 'dois' });
    const before = cache.stats.misses;
    layout({ a: 'outro', b: 'dois' });
    expect(cache.stats.misses - before).toBe(1);
  });
});
