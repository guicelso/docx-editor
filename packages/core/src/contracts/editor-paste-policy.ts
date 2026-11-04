import type { PastePolicy } from '@docx-editor.dev/core/store';

export type { PasteLanding, PastePolicy, PastedBlocks } from '@docx-editor.dev/core/store';

/** The host's say over rich pastes. @public */
export interface EditorPastePolicy {
  /**
   * Ask `policy` about every paste of the engine's own fragment or of external HTML: the blocks
   * that land, and what the paste would publish. Plain text, remote edits, undo and redo are not
   * asked. `null` removes it. The editor keeps the policy across `load`, `attach` and `detach`,
   * and hands it to each document it mounts.
   */
  setPastePolicy(policy: PastePolicy | null): void;
}
