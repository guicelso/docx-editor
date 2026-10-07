// Installing view-only content-control tags on a mounted surface: they paint, the token is the
// cache key, and nothing about them reaches the document.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import type { ContentControlTagDisplay } from '../../layout/content-control-tags.ts';
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
