// The prompt an empty content control shows, in the host's language.

import type { ContentControlKind } from '../store/package/content-control-nodes.ts';
import type { PaginatedSurfaceOptions } from './paginated-surface-options.ts';

/**
 * The prompt an empty control of `type` shows, from the host's translation. A translation
 * that does not know the key answers with the key (or nothing), and Word's default stands.
 */
export function placeholderPromptOf(
  translate: PaginatedSurfaceOptions['translate'],
  type: ContentControlKind
): string | undefined {
  const key = PLACEHOLDER_PROMPT_KEYS[type];
  if (!key || !translate) return undefined;
  const text = translate(key);
  return text && text !== key ? text : undefined;
}

/** The translation key of each prompt Word writes; a checkbox and a picture show none. */
const PLACEHOLDER_PROMPT_KEYS: Partial<Record<ContentControlKind, string>> = {
  richText: 'contentControl.prompt.text',
  plainText: 'contentControl.prompt.text',
  untyped: 'contentControl.prompt.text',
  comboBox: 'contentControl.prompt.list',
  dropDownList: 'contentControl.prompt.list',
  date: 'contentControl.prompt.date',
  docPartList: 'contentControl.prompt.buildingBlock',
};
