// The tags of a BLOCK-level content control in a real browser: one at the start of its first
// paragraph and one at the end of its last, as tall as the text beside them, and the right half of
// the opening tag types at the start of the control.

import { readOoxmlPackage } from '../packages/core/src/store/package/ooxml-package';
import type { OoxmlNode } from '../packages/core/src/store/package/ooxml-tree';
import {
  contentControlContentChildren,
  isContentControl,
} from '../packages/core/src/store/package/content-control-walk';
import { contentControlSubjectOf } from '../packages/core/src/layout/content-control-properties';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { Page } from '@playwright/test';
import { expect, test } from '@playwright/test';

const origin = process.env.CC_TAGS_ORIGIN ?? 'http://localhost:5273';

test.beforeEach(async ({ page }) => {
  await page.goto(`${origin}/?fixture=block-content-control-tags.docx&e2e=1`);
  await page.waitForFunction(() => window.__DOCX_EDITOR_E2E__?.fontMeasurer() === 'shaped');
  await page.evaluate(() => {
    const editor = window.__DOCX_EDITOR_E2E__!.getEditor() as DocxEditorInstance;
    editor.setContentControlTags({
      labelsOf: ({ tag }) => ({ open: { text: `${tag} ▸` }, close: { text: `◂ ${tag}` } }),
    });
  });
  await expect(page.locator('[data-cc-tag-control]')).toHaveCount(2);
});

/** The saved body with each control bracketed by its tag and `|` between paragraphs. */
async function savedText(page: Page): Promise<string> {
  const saved = await page.evaluate(async () =>
    Array.from((await window.__DOCX_EDITOR_E2E__!.saveBytes())!)
  );
  const opened = readOoxmlPackage(new Uint8Array(saved));
  if (!opened.ok) throw new Error(opened.reason);
  const walk = (nodes: readonly OoxmlNode[]): string =>
    nodes
      .map((node) => {
        if (node.kind === 'textValue') return node.value;
        if (isContentControl(node)) {
          return `${contentControlSubjectOf(node).tag}{${walk(contentControlContentChildren(node))}}`;
        }
        return node.kind === 'paragraph' ? `${walk(node.children)}|` : walk(node.children);
      })
      .join('');
  return walk([opened.package.parts.get(opened.package.mainDocumentPart)!.root]);
}

test('the opening tag stands on the first paragraph and the closing tag on the last', async ({
  page,
}) => {
  const lines = await page.evaluate(() =>
    Array.from(document.querySelectorAll<HTMLElement>('[data-cc-tag-control]')).map((chip) => ({
      edge: chip.dataset.ccTagEdge,
      line: (chip.closest('.layout-line')?.textContent ?? '').replace(/ /g, ' '),
    }))
  );
  expect(lines).toEqual([
    { edge: 'open', line: 'block-1 ▸first line' },
    { edge: 'close', line: 'second line◂ block-1' },
  ]);
});

test('every block tag is as tall as the text beside it', async ({ page }) => {
  const chips = await page.evaluate(() =>
    Array.from(document.querySelectorAll<HTMLElement>('[data-cc-tag-control]')).map((chip) => {
      const text = Array.from(
        chip.closest('.layout-line')?.querySelectorAll<HTMLElement>('.layout-run[data-start]') ?? []
      ).find((run) => !run.dataset.ccTagControl && (run.textContent ?? '').trim().length > 0);
      return {
        chip: chip.getBoundingClientRect().height,
        text: text?.getBoundingClientRect().height ?? Number.NaN,
      };
    })
  );
  for (const chip of chips) expect(Math.abs(chip.chip - chip.text)).toBeLessThan(0.5);
});

test('the right half of the opening tag types at the start of the control', async ({ page }) => {
  const chip = page.locator('[data-cc-tag-edge="open"]');
  const box = await chip.boundingBox();
  if (!box) throw new Error('no opening chip');
  await page.mouse.click(box.x + box.width * 0.8, box.y + box.height / 2);
  await page.keyboard.type('Y');
  expect(await savedText(page)).toContain('Before.|block-1{Yfirst line|second line|}After.|');
});

test('the left half of the opening tag types in a new paragraph before the control', async ({
  page,
}) => {
  const chip = page.locator('[data-cc-tag-edge="open"]');
  const box = await chip.boundingBox();
  if (!box) throw new Error('no opening chip');
  await page.mouse.click(box.x + box.width * 0.2, box.y + box.height / 2);
  await page.keyboard.type('XY');
  expect(await savedText(page)).toContain('Before.|XY|block-1{first line|second line|}After.|');
});

test('Enter on the right half of the closing tag opens a paragraph after the control', async ({
  page,
}) => {
  const chip = page.locator('[data-cc-tag-edge="close"]');
  const box = await chip.boundingBox();
  if (!box) throw new Error('no closing chip');
  await page.mouse.click(box.x + box.width * 0.8, box.y + box.height / 2);
  await page.keyboard.press('Enter');
  await page.keyboard.type('Z');
  expect(await savedText(page)).toContain('block-1{first line|second line|}Z|After.|');
});
