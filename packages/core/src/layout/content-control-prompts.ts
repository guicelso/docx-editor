// The host's text in a content control that holds its placeholder.
//
// The stored placeholder keeps its model range. The walk lays it out as one projected piece that
// paints the host's text in the placeholder's style and carries the stored text as `modelText`,
// so the text layout reads back, the clipboard and every offset still agree with the store.

import type { OoxmlElement } from '@docx-editor.dev/core/store';
import type { ContentControlPromptDisplay } from '../contracts/editor-content-control-view.ts';
import type { BlockControlEdges } from '../store/store/block-control-edges.ts';
import { contentControlPropertiesOf as storedPropertiesOf } from '../store/package/content-control-nodes.ts';
import { contentControlSubjectOf } from './content-control-properties.ts';
import type { FieldAwarePiece } from './field-pieces.ts';
import { plainTextFrom } from './plain-text-pieces.ts';

/** The host's text for a control showing its placeholder, or null to show the stored one. */
export function contentControlPromptOf(
  control: OoxmlElement,
  prompts: ContentControlPromptDisplay | undefined
): string | null {
  if (!prompts || !storedPropertiesOf(control).showingPlaceholder) return null;
  return prompts.promptOf(contentControlSubjectOf(control)) || null;
}

/**
 * The host's text for the block control whose only paragraph this is, innermost first. A block
 * control holding several paragraphs shows its stored placeholder: one projected piece cannot
 * stand for text that several paragraphs hold.
 */
export function blockContentControlPromptOf(
  edges: BlockControlEdges,
  prompts: ContentControlPromptDisplay | undefined
): string | null {
  if (!prompts) return null;
  for (let index = edges.opens.length - 1; index >= 0; index -= 1) {
    const control = edges.opens[index]!;
    if (!edges.closes.includes(control)) continue;
    const prompt = contentControlPromptOf(control, prompts);
    if (prompt !== null) return prompt;
  }
  return null;
}

/**
 * Lay the stored placeholder, `pieces` from `from` on, out as one piece showing `prompt`.
 *
 * Only contiguous plain text qualifies (`plainTextFrom`): a placeholder that is tracked, linked,
 * holds a field, a drawing, a break or another control's tag keeps its stored pieces: the view
 * never hides what the file marks.
 */
export function projectContentControlPrompt(
  pieces: FieldAwarePiece[],
  from: number,
  prompt: string
): void {
  const plain = plainTextFrom(pieces, from);
  if (!plain) return;
  pieces.splice(from, pieces.length - from, {
    text: prompt,
    props: plain.first.props,
    style: plain.first.style,
    start: plain.first.start,
    end: plain.end,
    projected: true,
    modelText: plain.text,
  });
}
