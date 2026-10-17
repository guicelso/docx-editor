import type { WriteAdmission } from '@docx-editor.dev/core/store';

export type { WriteAdmission } from '@docx-editor.dev/core/store';

/** The host's say over what the editor writes. @public */
export interface EditorWriteAdmission {
  /**
   * Ask `admission` about every part a local write would change, as the part would be
   * published; a string refuses the whole write, with that reason in `lastRejection`. Remote
   * edits, undo and redo are not asked. `null` removes it. The editor keeps the admission across
   * `load`, `attach` and `detach`, and hands it to each document it mounts.
   */
  setWriteAdmission(admission: WriteAdmission | null): void;
}
