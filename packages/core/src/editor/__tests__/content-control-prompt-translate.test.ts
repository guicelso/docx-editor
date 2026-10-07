// The prompt an empty content control shows comes from the host's translation, live: on open,
// and whenever an edit creates or empties a control.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import {
  contentControlPropertiesOf,
  contentControlTextOf,
  contentControlsIn,
} from '../../store/index.ts';
import { mountPaginatedSurface, type PaginatedSurface } from '../paginated-surface.ts';
import { docx } from './paginated-surface-fixtures.ts';

const EMPTY =
  '<w:p><w:r><w:t xml:space="preserve">a </w:t></w:r>' +
  '<w:sdt><w:sdtPr><w:tag w:val="t"/><w:richText/></w:sdtPr><w:sdtContent/></w:sdt></w:p>';

const PT: Record<string, string> = { 'contentControl.prompt.text': 'digite o trecho' };

const mounted: PaginatedSurface[] = [];
afterEach(() => {
  for (const surface of mounted.splice(0)) surface.destroy();
});

function open(translate?: (key: string) => string): PaginatedSurface {
  const result = mountPaginatedSurface(document.createElement('div'), docx(EMPTY), {
    scale: 1,
    ...(translate ? { translate } : {}),
  });
  if (!result.ok) throw new Error(result.reason);
  mounted.push(result.surface);
  return result.surface;
}

const promptOf = (surface: PaginatedSurface) =>
  contentControlTextOf(contentControlsIn(surface.session.part().root)[0]!.node);

describe('the placeholder prompt follows the host translation', () => {
  test('an empty control opens showing the translated prompt', () => {
    expect(promptOf(open((key) => PT[key] ?? key))).toBe('digite o trecho');
  });

  test('a translation that does not know the key leaves Word’s default', () => {
    expect(promptOf(open((key) => key))).toBe('Click here to enter text.');
    expect(promptOf(open())).toBe('Click here to enter text.');
  });

  test('an edit that creates an empty control asks the translation in force now', () => {
    let language: Record<string, string> = {};
    const surface = open((key) => language[key] ?? key);
    language = PT;
    surface.setTranslate((key) => language[key] ?? key);
    const paragraphId = surface.session.paragraphIds()[0]!;
    const result = surface.applyAutomationOps(() => [
      {
        op: 'insertContentControl',
        paragraphId,
        start: 0,
        end: 0,
        type: 'richText',
        tag: 'new',
      },
    ]);
    expect(result.committed).toBe(true);
    const created = contentControlsIn(surface.session.part().root).find(
      (entry) => contentControlPropertiesOf(entry.node).tag === 'new'
    );
    expect(created && contentControlTextOf(created.node)).toBe('digite o trecho');
  });
});
