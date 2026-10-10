// The host's view is part of each paragraph's cache identity, never of the whole document's: a
// changed answer measures again only the paragraphs that show it.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, type OoxmlPart } from '../../store/package/ooxml-tree.ts';
import type { ContentControlTagDisplay } from '../../contracts/editor-content-control-view.ts';
import { createParagraphLayoutCache } from '../layout-cache.ts';
import { createLayoutSession } from '../layout-session.ts';
import type { PendingLine } from '../paragraph-flow.ts';
import { createFixedMeasurer, layoutSemanticDocument } from '../semantic-layout.ts';
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
const paragraph = (inner: string) => `<w:p>${inner}</w:p>`;
const inline = (tag: string, inner: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

/** Three paragraphs: one per tagged control, and one with no control at all. */
const PART = documentOf(
  paragraph(run('um ') + inline('a', run('A'))) +
    paragraph(run('dois ') + inline('b', run('B'))) +
    paragraph(run('tres'))
);

const tags = (labels: Readonly<Record<string, string>>): ContentControlView => ({
  tags: {
    labelsOf: ({ tag }) =>
      tag === undefined ? null : { open: { text: `${labels[tag]} ▸` }, close: { text: '◂' } },
  } satisfies ContentControlTagDisplay,
});

/** A layout run that keeps its break cache and its session, as the surface does. */
function passes() {
  const cache = createParagraphLayoutCache<readonly PendingLine[]>();
  const session = createLayoutSession();
  const measurer = createFixedMeasurer(6, 14);
  return {
    cache,
    session,
    layout: (view: ContentControlView | undefined) =>
      layoutSemanticDocument(PART, 0, {
        measurer,
        cache,
        session,
        ...(view ? { contentControlView: view } : {}),
      }),
  };
}

describe('the view keys each paragraph by its own answers', () => {
  test('a changed label breaks only the paragraph that shows it', () => {
    const run = passes();
    run.layout(tags({ a: 'A1', b: 'B' }));
    const before = run.cache.stats.misses;
    run.layout(tags({ a: 'A2', b: 'B' }));
    expect(run.cache.stats.misses - before).toBe(1);
  });

  test('a new view with the same answers breaks nothing again', () => {
    const run = passes();
    const first = run.layout(tags({ a: 'A', b: 'B' }));
    const before = run.cache.stats.misses;
    const second = run.layout(tags({ a: 'A', b: 'B' }));
    expect(run.cache.stats.misses).toBe(before);
    expect(second.pages).toBe(first.pages);
  });

  test('the same view object reuses every page', () => {
    const run = passes();
    const view = tags({ a: 'A', b: 'B' });
    const first = run.layout(view);
    expect(run.layout(view).pages).toBe(first.pages);
  });

  test('installing a view leaves a paragraph without a control on its cached break', () => {
    const run = passes();
    run.layout(undefined);
    const before = run.cache.stats.misses;
    run.layout(tags({ a: 'A', b: 'B' }));
    // The two paragraphs with a control break again; the third keeps its key.
    expect(run.cache.stats.misses - before).toBe(2);
  });
});
