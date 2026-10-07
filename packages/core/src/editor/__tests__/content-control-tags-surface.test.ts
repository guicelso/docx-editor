// Installing view-only content-control tags on a mounted surface: they paint, the token is the
// cache key, and nothing about them reaches the document.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import type { ContentControlTagDisplay } from '../../layout/content-control-tags.ts';
import { contentControlsInLayout } from '../../layout/semantic-interaction.ts';
import type { PaginatedSurface } from '../paginated-surface.ts';
import { mount } from './paginated-surface-fixtures.ts';

const BODY =
  '<w:p><w:r><w:t xml:space="preserve">CPF </w:t></w:r>' +
  '<w:sdt><w:sdtPr><w:tag w:val="t"/></w:sdtPr>' +
  '<w:sdtContent><w:r><w:t>RG</w:t></w:r></w:sdtContent></w:sdt></w:p>';

const tags = (token: string): ContentControlTagDisplay => ({
  token,
  labelsOf: () => ({ open: { text: `${token} ▸` }, close: { text: '◂' } }),
});

const mounted: PaginatedSurface[] = [];
afterEach(() => {
  for (const surface of mounted.splice(0)) surface.destroy();
});

function surfaceOf(): { surface: PaginatedSurface; container: HTMLElement } {
  const opened = mount(BODY);
  mounted.push(opened.surface);
  return opened;
}

const chipsIn = (container: HTMLElement) =>
  [...container.querySelectorAll<HTMLElement>('[data-cc-tag-control]')].map(
    (chip) => chip.textContent?.replace(/[  ]/g, ' ').trim() ?? ''
  );

const documentXmlOf = (surface: PaginatedSurface) =>
  strFromU8(unzipSync(surface.save())['word/document.xml']!);

describe('content-control tags on the surface', () => {
  test('installing a display paints the chips, and null takes them away', () => {
    const { surface, container } = surfaceOf();
    surface.setContentControlTags(tags('A'));
    expect(chipsIn(container)).toEqual(['A ▸', '◂']);
    surface.setContentControlTags(null);
    expect(chipsIn(container)).toEqual([]);
  });

  test('the token already installed lays out nothing; a new token does', () => {
    const { surface, container } = surfaceOf();
    surface.setContentControlTags(tags('A'));
    const laidOut = surface.layout();
    surface.setContentControlTags(tags('A'));
    expect(surface.layout()).toBe(laidOut);
    surface.setContentControlTags(tags('B'));
    expect(surface.layout()).not.toBe(laidOut);
    expect(chipsIn(container)).toEqual(['B ▸', '◂']);
  });

  test('the saved document is the same with the tags on or off', () => {
    const { surface } = surfaceOf();
    const untagged = documentXmlOf(surface);
    surface.setContentControlTags(tags('A'));
    expect(documentXmlOf(surface)).toBe(untagged);
  });

  test('the selected text never carries a tag', () => {
    const { surface } = surfaceOf();
    surface.setContentControlTags(tags('A'));
    surface.selectAll();
    expect(surface.selectedText()).toBe('CPF RG');
  });
});

describe('field tones on the surface', () => {
  const FIELD =
    '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> MERGEFIELD "field:a" </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>«cpf»</w:t></w:r>' +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';

  test('installing tones repaints the fields without laying anything out', () => {
    const opened = mount(FIELD);
    mounted.push(opened.surface);
    const laidOut = opened.surface.layout();
    opened.surface.setFieldTones((instruction) =>
      instruction.includes('"field:') ? 'variable' : undefined
    );
    expect(opened.surface.layout()).toBe(laidOut);
    expect(
      opened.container.querySelector<HTMLElement>('[data-field-atom]')?.dataset.fieldTone
    ).toBe('variable');
    opened.surface.setFieldTones(null);
    expect(
      opened.container.querySelector<HTMLElement>('[data-field-atom]')?.dataset.fieldTone
    ).toBeUndefined();
  });
});

describe('hovering a content control', () => {
  test('the control under a resting pointer is published and its chrome retinted', () => {
    const opened = mount(BODY);
    mounted.push(opened.surface);
    const { surface, container } = opened;
    const layout = surface.layout();
    const page = layout.pages[0]!;
    const fragment = contentControlsInLayout(layout)[0]!.fragments[0]!;
    const target = container.querySelector<HTMLElement>('[data-paragraph-id]')!;
    const move = (x: number, y: number) =>
      target.dispatchEvent(
        new PointerEvent('pointermove', { bubbles: true, clientX: x, clientY: y })
      );

    move(
      page.contentBox.x + fragment.box.x + fragment.box.width / 2,
      page.contentBox.y + fragment.box.y + fragment.box.height / 2
    );
    const hovered = surface.state().contentControls.hoveredControlId;
    expect(hovered).not.toBeNull();
    const chrome = container.querySelector<HTMLElement>(
      `.docx-content-control-chrome[data-docx-content-control="${hovered}"]`
    );
    expect(chrome?.dataset.hover).toBe('');
    expect(surface.layout()).toBe(layout);

    container
      .querySelector('[data-paragraph-id]')!
      .closest('[contenteditable="true"]')
      ?.dispatchEvent(new PointerEvent('pointerleave'));
    expect(surface.state().contentControls.hoveredControlId).toBeNull();
    expect(chrome?.dataset.hover).toBeUndefined();
  });
});
