/** The host's write admission, kept by the facade across surfaces. */

import type { WriteAdmission } from '@docx-editor.dev/core/store';
import type { EditorWriteAdmission } from '../contracts/editor-write-admission.ts';
import type { PaginatedSurface, PaginatedSurfaceOptions } from './paginated-surface.ts';

/**
 * The admission and its setter. It answers for the host's structure, never for one document,
 * so every surface the editor rebuilds mounts with it, and the live one takes every change.
 */
export function createEditorWriteAdmission(host: { surface(): PaginatedSurface | null }): {
  readonly mountOptions: () => Pick<PaginatedSurfaceOptions, 'writeAdmission'>;
  readonly members: EditorWriteAdmission;
} {
  let admission: WriteAdmission | null = null;
  return {
    mountOptions: () => (admission ? { writeAdmission: admission } : {}),
    members: {
      setWriteAdmission(next) {
        admission = next;
        host.surface()?.setWriteAdmission(next);
      },
    },
  };
}
