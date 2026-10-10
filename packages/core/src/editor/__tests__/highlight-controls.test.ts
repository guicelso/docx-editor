// A highlight that names a content control marks the control whole: its content in every
// paragraph it reaches and the tags at its edges and its children's — the area its boundary
// outlines — whatever level the control sits at.

import { afterEach, describe, expect, test } from 'bun:test';
import { contentControlPropertiesOf } from '../../store/package/content-control-nodes.ts';
import type { OoxmlNode } from '../../store/package/ooxml-tree.ts';
import type { LayoutBox, SemanticLayout } from '../../layout/semantic-records.ts';
import type { DocxEditorInstance } from '../docx-editor-types.ts';
import { docx } from './paginated-surface-fixtures.ts';
import { mountAnchorEditor } from './scroll-to-anchor-fixture.ts';

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const paragraph = (inner: string) => `<w:p>${inner}</w:p>`;
const control = (tag: string, inner: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w:richText/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

/** `Antes | B[ um | dois I{três} ] | Depois`. */
const BODY =
  paragraph(run('Antes')) +
  control('B', paragraph(run('um')) + paragraph(run('dois ') + control('I', run('três')))) +
  paragraph(run('Depois'));

const cleanups: (() => void)[] = [];
afterEach(() => {
  for (const cleanup of cleanups.splice(0).reverse()) cleanup();
});

function mount(): { editor: DocxEditorInstance; host: HTMLElement } {
  const mounted = mountAnchorEditor(docx(BODY));
  cleanups.push(mounted.destroy);
  mounted.editor.surface!.setContentControlTags({
    token: 'veil',
    labelsOf: ({ tag }) => ({ open: { text: `${tag}▸` }, close: { text: `◂${tag}` } }),
  });
  return mounted;
}

function controlId(editor: DocxEditorInstance, tag: string): string {
  let found: string | null = null;
  const walk = (node: OoxmlNode): void => {
    if (found || node.kind === 'textValue') return;
    if (node.kind === 'contentControl' && contentControlPropertiesOf(node).tag === tag)
      found = node.id;
    else node.children.forEach(walk);
  };
  walk(editor.surface!.session.part().root);
  if (found === null) throw new Error(`no control ${tag}`);
  return found;
}

const marks = (host: HTMLElement, name: string) => [
  ...host.querySelectorAll<HTMLElement>(`[data-highlight-set="${name}"] .docx-text-highlight`),
];

function inside(inner: LayoutBox, outer: LayoutBox): boolean {
  return (
    inner.x >= outer.x - 0.01 &&
    inner.y >= outer.y - 0.01 &&
    inner.x + inner.width <= outer.x + outer.width + 0.01 &&
    inner.y + inner.height <= outer.y + outer.height + 0.01
  );
}

/** Every tag span of a control and its children, with the page it is on. */
function tagBoxes(layout: SemanticLayout, ids: readonly string[]) {
  const boxes: { pageIndex: number; box: LayoutBox }[] = [];
  layout.pages.forEach((page, pageIndex) => {
    for (const fragment of page.fragments)
      for (const line of (
        fragment as {
          lines?: {
            spans: readonly import('../../layout/semantic-records.ts').StyleSpanRecord[];
          }[];
        }
      ).lines ?? [])
        for (const span of line.spans)
          if (span.contentControlTag && ids.includes(span.contentControlTag.controlId))
            boxes.push({ pageIndex, box: span.box });
  });
  return boxes;
}

describe('a highlight that names a content control', () => {
  test('marks a block control over its paragraphs and every tag it and its children draw', () => {
    const { editor, host } = mount();
    const block = controlId(editor, 'B');
    const inline = controlId(editor, 'I');

    expect(editor.setHighlights('veil', [{ controlId: block }], { blend: 'cover' })).toEqual({
      applied: 1,
      unavailable: 0,
    });

    const layout = editor.surface!.publishedLayout();
    const fragments = layout.contentControls!.find((record) => record.id === block)!.fragments;
    expect(marks(host, 'veil')).toHaveLength(fragments.length);
    for (const tag of tagBoxes(layout, [block, inline])) {
      expect(
        fragments.some(
          (fragment) => fragment.pageIndex === tag.pageIndex && inside(tag.box, fragment.box)
        )
      ).toBe(true);
    }
  });

  test('marks an inline control over its content and its own tags, line by line', () => {
    const { editor, host } = mount();
    const inline = controlId(editor, 'I');

    editor.setHighlights('veil', [{ controlId: inline }], { blend: 'cover' });

    const layout = editor.surface!.publishedLayout();
    const fragments = layout.contentControls!.find((record) => record.id === inline)!.fragments;
    expect(marks(host, 'veil')).toHaveLength(fragments.length);
    for (const tag of tagBoxes(layout, [inline])) {
      expect(fragments.some((fragment) => inside(tag.box, fragment.box))).toBe(true);
    }
  });

  test('a control and a text range share one set', () => {
    const { editor } = mount();
    const [match] = editor.findMatches('Depois');

    expect(
      editor.setHighlights('veil', [{ controlId: controlId(editor, 'I') }, match!], {
        blend: 'cover',
      })
    ).toEqual({ applied: 2, unavailable: 0 });
  });

  test('a control the document does not hold is unavailable, and one removed stops painting', () => {
    const { editor, host } = mount();
    const inline = controlId(editor, 'I');

    expect(editor.setHighlights('veil', [{ controlId: 'missing' }])).toEqual({
      applied: 0,
      unavailable: 1,
    });

    editor.setHighlights('veil', [{ controlId: inline }]);
    editor.surface!.applyAutomationOps(() => [
      { op: 'removeContentControl', controlId: inline, keepContent: true },
    ]);
    editor.surface!.repaintHighlights();
    expect(marks(host, 'veil')).toHaveLength(0);
  });

  test('the hit under a control mark names the control', () => {
    const { editor, host } = mount();
    const inline = controlId(editor, 'I');
    const target = { controlId: inline };
    editor.setHighlights('veil', [target], { blend: 'cover' });
    host.querySelector<HTMLElement>('.docx-text-highlight-overlay')!.getBoundingClientRect = () =>
      ({ left: 0, top: 0 }) as DOMRect;
    const [mark] = marks(host, 'veil');

    const [hit] = editor.getHighlightsAt<typeof target>(
      Number.parseFloat(mark!.style.left) + 1,
      Number.parseFloat(mark!.style.top) + 1
    );

    expect(hit?.controlId).toBe(inline);
    expect(hit?.range).toBe(target);
    expect(hit && 'start' in hit).toBe(false);
  });

  test('a target that names both a paragraph and a control is refused', () => {
    const { editor } = mount();

    expect(() =>
      editor.setHighlights('veil', [
        { controlId: controlId(editor, 'I'), blockId: 'x', start: 0, length: 1 } as never,
      ])
    ).toThrow(TypeError);
  });
});
