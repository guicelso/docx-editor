// The Vue twin of the composed-arrangement cases in `packages/react/test/toolbar-overflow-react.test.tsx`.
//
// A part child for a registry slot the default arrangement does not draw joins its registry
// group, so it collapses with it; only a child that drives no slot stays fixed.

import './dom-setup.ts';

import { afterEach, describe, expect, test } from 'bun:test';
import { h } from 'vue';
import { DocxEditorToolbar } from '../src/editor/toolbar';
import { docx, flush } from './helpers/fixtures';
import { mountEditorTree, type MountedEditor } from './helpers/mount';

let mounted: MountedEditor | undefined;
afterEach(() => {
  mounted?.unmount();
  mounted = undefined;
});

const composed = () =>
  h(
    DocxEditorToolbar,
    {},
    {
      default: () => [
        h(DocxEditorToolbar.TableInsert),
        h(DocxEditorToolbar.Button, { slot: 'insert.pageBreak' }),
        h(DocxEditorToolbar.Button, { slot: 'review.paragraphMarks' }),
        h(DocxEditorToolbar.Button, { slot: 'contentControl.inspector', hidden: true }),
        h(DocxEditorToolbar.Action, { label: 'Host action', onSelect: () => {} }),
      ],
    }
  );

describe('the Vue toolbar arranges composed registry parts', () => {
  test('each joins its registry group; only the child that drives no slot stays fixed', async () => {
    mounted = mountEditorTree(composed, docx('<w:p><w:r><w:t>alpha</w:t></w:r></w:p>'));
    await flush();
    const toolbar = mounted.container;
    const groupOf = (slot: string) =>
      toolbar.querySelector(`[data-slot="${slot}"]`)?.closest('.docx-toolbar__group');

    expect(groupOf('table.insert')?.getAttribute('data-toolbar-group')).toBe('table');
    expect(groupOf('insert.pageBreak')?.getAttribute('data-toolbar-group')).toBe('insert');
    expect(groupOf('review.paragraphMarks')).toBe(groupOf('review.comments'));

    const host = toolbar
      .querySelector('button[aria-label="Host action"]')
      ?.closest('.docx-toolbar__group');
    expect(host?.hasAttribute('data-toolbar-fixed')).toBe(true);
    expect(host?.querySelector('[data-slot]')).toBeNull();
  });

  test('a hidden part for a slot the arrangement does not draw adds no group', async () => {
    mounted = mountEditorTree(composed, docx('<w:p><w:r><w:t>alpha</w:t></w:r></w:p>'));
    await flush();

    expect(mounted.container.querySelector('[data-slot="contentControl.inspector"]')).toBeNull();
    expect(mounted.container.querySelector('[data-toolbar-group="contentControl"]')).toBeNull();
  });
});
