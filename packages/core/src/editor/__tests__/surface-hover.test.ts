// What a resting pointer is on: one answer from the layout hit test, published on its own channel
// and read by the chrome retint, never through the selection.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import type { EditorHover } from '../../contracts/editor.ts';
import type { LayoutBox, PageRecord, SemanticLayout } from '../../layout/semantic-records.ts';
import { contentControlsInLayout } from '../../layout/semantic-interaction.ts';
import { contentControlWidgetArea } from '../../output/semantic-paint-content-controls.ts';
import { mountPaginatedSurface, type PaginatedSurface } from '../paginated-surface.ts';
import { createPublishSignal } from '../surface-publish-signal.ts';
import { docx } from './paginated-surface-fixtures.ts';
import { mountAnchorEditor } from './scroll-to-anchor-fixture.ts';

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const control = (tag: string, properties: string, inner: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/>${properties}</w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
const field = (instruction: string, result: string) =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> MERGEFIELD "${instruction}" </w:instrText></w:r>` +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${run(result)}` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

const CONTROL = `<w:p>${run('CPF ')}${control('t', '', run('RG'))}</w:p>`;
const FIELD_IN_CONTROL = `<w:p>${run('um ')}${control('t', '', field('field:a', '«cpf»'))}${run(' fim')}</w:p>`;
const DROPDOWN = `<w:p>${run('Estado ')}${control(
  'd',
  '<w:dropDownList><w:listItem w:displayText="SC" w:value="SC"/></w:dropDownList>',
  run('SC')
)}</w:p>`;

const mounted: PaginatedSurface[] = [];
const cleanups: (() => void)[] = [];
afterEach(() => {
  for (const surface of mounted.splice(0)) surface.destroy();
  for (const cleanup of cleanups.splice(0).reverse()) cleanup();
});

function surfaceOf(body: string) {
  const hovers: EditorHover[] = [];
  const container = document.createElement('div');
  const opened = mountPaginatedSurface(container, docx(body), {
    scale: 1,
    onHoverChange: (hover) => hovers.push(hover),
  });
  if (!opened.ok) throw new Error(opened.reason);
  mounted.push(opened.surface);
  return { surface: opened.surface, container, hovers };
}

/** happy-dom lays nothing out: the pages sit at the viewport origin, at scale 1. */
function moveTo(container: HTMLElement, x: number, y: number): void {
  container
    .querySelector('[data-paragraph-id]')!
    .dispatchEvent(new PointerEvent('pointermove', { bubbles: true, clientX: x, clientY: y }));
}

function leave(container: HTMLElement): void {
  container
    .querySelector('[data-paragraph-id]')!
    .closest('[contenteditable]')
    ?.dispatchEvent(new PointerEvent('pointerleave'));
}

const centre = (page: PageRecord, box: LayoutBox) => ({
  x: page.contentBox.x + box.x + box.width / 2,
  y: page.contentBox.y + box.y + box.height / 2,
});

const onScreen = (page: PageRecord, box: LayoutBox) => ({
  x: page.contentBox.x + box.x,
  y: page.contentBox.y + box.y,
  width: box.width,
  height: box.height,
});

function fieldSpanOf(layout: SemanticLayout): LayoutBox {
  for (const fragment of layout.pages[0]!.fragments) {
    if (fragment.kind !== 'paragraph') continue;
    const span = fragment.lines
      .flatMap((line) => line.spans)
      .find((candidate) => candidate.fieldAtom);
    if (span) return span.box;
  }
  throw new Error('no field painted');
}

function fieldIdOf(layout: SemanticLayout): string | undefined {
  for (const fragment of layout.pages[0]!.fragments) {
    if (fragment.kind !== 'paragraph') continue;
    const span = fragment.lines
      .flatMap((line) => line.spans)
      .find((candidate) => candidate.fieldAtom);
    if (span) return span.fieldAtom?.field?.fieldId;
  }
  return undefined;
}

describe('the hover on the surface', () => {
  test('the control under a resting pointer is published apart from the selection, with its box', () => {
    const { surface, container, hovers } = surfaceOf(CONTROL);
    const layout = surface.layout();
    const page = layout.pages[0]!;
    const record = contentControlsInLayout(layout)[0]!;
    const fragment = record.fragments[0]!.box;
    const signal = createPublishSignal();
    signal.adopt(surface);

    const at = centre(page, fragment);
    moveTo(container, at.x, at.y);

    expect(hovers).toEqual([{ controlId: record.id, field: null, box: onScreen(page, fragment) }]);
    expect(signal.moved(surface.state(), surface)).toBe(false);
    const chrome = container.querySelector<HTMLElement>(
      `.docx-content-control-chrome[data-docx-content-control="${record.id}"]`
    );
    expect(chrome?.dataset.hover).toBe('');
    expect(surface.layout()).toBe(layout);

    leave(container);
    expect(hovers.at(-1)).toEqual({ controlId: null, field: null, box: null });
    expect(chrome?.dataset.hover).toBeUndefined();
  });

  test('moving inside the same piece publishes nothing more', () => {
    const { surface, container, hovers } = surfaceOf(CONTROL);
    const page = surface.layout().pages[0]!;
    const fragment = contentControlsInLayout(surface.layout())[0]!.fragments[0]!.box;
    const at = centre(page, fragment);

    moveTo(container, at.x, at.y);
    moveTo(container, at.x + 1, at.y);

    expect(hovers).toHaveLength(1);
  });

  test('a field is hovered on its glyphs, inside its control, with the span it is painted in', () => {
    const { surface, container, hovers } = surfaceOf(FIELD_IN_CONTROL);
    const layout = surface.layout();
    const page = layout.pages[0]!;
    const span = fieldSpanOf(layout);
    const at = centre(page, span);

    moveTo(container, at.x, at.y);

    const hover = hovers.at(-1)!;
    expect(hover.controlId).toBe(contentControlsInLayout(layout)[0]!.id);
    expect(hover.field?.instruction).toContain('"field:a"');
    expect(hover.field?.fieldId).toBe(fieldIdOf(layout));
    expect(hover.box).toEqual(onScreen(page, span));
  });

  test('a field beside the pointer is not hovered', () => {
    const { surface, container, hovers } = surfaceOf(FIELD_IN_CONTROL);
    const page = surface.layout().pages[0]!;
    const span = fieldSpanOf(surface.layout());
    const on = centre(page, span);
    moveTo(container, on.x, on.y);

    moveTo(container, on.x, page.contentBox.y + span.y + span.height * 4);

    expect(hovers.at(-2)?.field).not.toBeNull();
    expect(hovers.at(-1)?.field).toBeNull();
  });

  test('a control keeps the hover across its value widget, painted outside it', () => {
    const { surface, container, hovers } = surfaceOf(DROPDOWN);
    const layout = surface.layout();
    const page = layout.pages[0]!;
    const record = contentControlsInLayout(layout)[0]!;
    const widget = contentControlWidgetArea(page, record)!;
    const at = centre(page, widget);

    moveTo(container, at.x, at.y);

    expect(hovers.at(-1)?.controlId).toBe(record.id);
  });
});

describe('the hover on the editor', () => {
  test('hosts hear it as hoverChange and never as selectionChange, and a load ends it', () => {
    const opened = mountAnchorEditor(docx(CONTROL));
    cleanups.push(opened.destroy);
    const { editor, host } = opened;
    const hovers: EditorHover[] = [];
    let selectionChanges = 0;
    editor.on('hoverChange', (hover) => hovers.push(hover));
    editor.on('selectionChange', () => {
      selectionChanges += 1;
    });
    const layout = editor.surface!.layout();
    const page = layout.pages[0]!;
    const pages = host.querySelector<HTMLElement>('.docx-pages')!;
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
    const scale = editor.getRenderScale();
    const box = contentControlsInLayout(layout)[0]!.fragments[0]!.box;
    const at = centre(page, box);

    pages.dispatchEvent(
      new PointerEvent('pointermove', {
        bubbles: true,
        clientX: at.x * scale,
        clientY: at.y * scale,
      })
    );
    expect(hovers.at(-1)?.controlId).not.toBeNull();
    expect(selectionChanges).toBe(0);

    editor.load(docx(CONTROL));
    expect(hovers.at(-1)).toEqual({ controlId: null, field: null, box: null });
  });
});
