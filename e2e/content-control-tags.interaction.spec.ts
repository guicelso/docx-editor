// Content-control tags in a real browser: the half of a chip that is clicked decides where the
// typed text lands — in front of a control, at its start, or after it.

import { readOoxmlPackage } from '../packages/core/src/store/package/ooxml-package';
import type { OoxmlNode } from '../packages/core/src/store/package/ooxml-tree';
import {
  contentControlContentChildren,
  isContentControl,
} from '../packages/core/src/store/package/content-control-walk';
import { contentControlTagSubjectOf } from '../packages/core/src/layout/content-control-tags';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { Page } from '@playwright/test';
import { expect, test } from '@playwright/test';

const origin = process.env.CC_TAGS_ORIGIN ?? 'http://localhost:5273';

test.beforeEach(async ({ page }) => {
  await page.goto(`${origin}/?fixture=content-control-tags.docx&e2e=1`);
  await page.waitForFunction(() => window.__DOCX_EDITOR_E2E__?.fontMeasurer() === 'shaped');
  await page.evaluate(() => {
    const surface = (window.__DOCX_EDITOR_E2E__!.getEditor() as DocxEditorInstance).surface;
    if (!surface) throw new Error('the editor has no surface yet');
    surface.setContentControlTags({
      token: 'tag-names',
      labelsOf: ({ tag }) => ({ open: { text: `${tag} ▸` }, close: { text: `◂ ${tag}` } }),
    });
  });
  await expect(page.locator('[data-cc-tag-control]')).toHaveCount(12);
});

async function clickChip(
  page: Page,
  tag: string,
  edge: 'open' | 'close',
  half: 'left' | 'right'
): Promise<void> {
  const chip = page.locator(`[data-cc-tag-edge="${edge}"]`).filter({ hasText: tag });
  const box = await chip.boundingBox();
  if (!box) throw new Error(`no ${edge} chip for ${tag}`);
  await page.mouse.click(box.x + box.width * (half === 'left' ? 0.2 : 0.8), box.y + box.height / 2);
}

/** The saved paragraph with each control bracketed by its tag: `Status: group-1{case-1{alpha}…}`. */
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
          return `${contentControlTagSubjectOf(node).tag}{${walk(contentControlContentChildren(node))}}`;
        }
        return walk(node.children);
      })
      .join('');
  return walk([opened.package.parts.get(opened.package.mainDocumentPart)!.root]);
}

test('the right half of an opening chip types at the start of its control', async ({ page }) => {
  await clickChip(page, 'fallback-1', 'open', 'right');
  await page.keyboard.type('Y');
  expect(await savedText(page)).toContain('fallback-1{Ybeta}');
});

test('the left half of an opening chip types in front of its control, letter after letter', async ({
  page,
}) => {
  await clickChip(page, 'group-2', 'open', 'left');
  await page.keyboard.type('AB');
  expect(await savedText(page)).toContain('. Then: ABgroup-2{');
});

test('the right half of a closing chip types after its control', async ({ page }) => {
  await clickChip(page, 'group-1', 'close', 'right');
  await page.keyboard.type('Z');
  expect(await savedText(page)).toContain('{beta}}Z. Then: ');
});

test('typing at the end of a control stays in it, keystroke after keystroke', async ({ page }) => {
  await clickChip(page, 'case-1', 'close', 'left');
  await page.keyboard.type('XY');
  await page.keyboard.press('Backspace');
  await page.keyboard.type('Z');
  expect(await savedText(page)).toContain('case-1{alphaXZ}fallback-1{beta}');
});

test('every chip draws its label centred, with the same room on both sides and nothing overflowing', async ({
  page,
}) => {
  const rooms = await page.evaluate(() =>
    Array.from(document.querySelectorAll<HTMLElement>('[data-cc-tag-control]')).map((chip) => {
      const label = document.createRange();
      label.selectNodeContents(chip);
      const ink = label.getBoundingClientRect();
      const box = chip.getBoundingClientRect();
      return {
        left: ink.left - box.left,
        right: box.right - ink.right,
        overflow: chip.scrollWidth - chip.clientWidth,
      };
    })
  );
  for (const room of rooms) {
    expect(room.overflow).toBe(0);
    expect(room.left).toBeGreaterThan(3);
    expect(Math.abs(room.left - room.right)).toBeLessThan(0.5);
  }
});

test('every chip is as tall as the text beside it, and its fill keeps its rounded corners', async ({
  page,
}) => {
  const chips = await page.evaluate(() =>
    Array.from(document.querySelectorAll<HTMLElement>('[data-cc-tag-control]')).map((chip) => {
      const text = Array.from(
        chip.closest('.layout-line')?.querySelectorAll<HTMLElement>('.layout-run[data-start]') ?? []
      ).find((run) => !run.dataset.ccTagControl && (run.textContent ?? '').trim().length > 0);
      const style = getComputedStyle(chip);
      return {
        chip: chip.getBoundingClientRect().height,
        text: text?.getBoundingClientRect().height ?? Number.NaN,
        border: parseFloat(style.borderLeftWidth) + parseFloat(style.borderRightWidth),
        radius: parseFloat(style.borderTopLeftRadius),
      };
    })
  );
  for (const chip of chips) {
    expect(Math.abs(chip.chip - chip.text)).toBeLessThan(0.5);
    expect(chip.border).toBe(0);
    expect(chip.radius).toBeGreaterThan(0);
  }
});
