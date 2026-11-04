/** The host's policy over rich pastes, kept by the facade across surfaces. */

import type { PastePolicy } from '@docx-editor.dev/core/store';
import type { EditorPastePolicy } from '../contracts/editor-paste-policy.ts';
import type { PaginatedSurface, PaginatedSurfaceOptions } from './paginated-surface.ts';

/**
 * The policy and its setter. It answers for the host's structure, never for one document, so
 * every surface the editor rebuilds mounts with it, and the live one takes every change.
 */
export function createEditorPastePolicy(host: { surface(): PaginatedSurface | null }): {
  readonly mountOptions: () => Pick<PaginatedSurfaceOptions, 'pastePolicy'>;
  readonly members: EditorPastePolicy;
} {
  let policy: PastePolicy | null = null;
  return {
    mountOptions: () => (policy ? { pastePolicy: policy } : {}),
    members: {
      setPastePolicy(next) {
        policy = next;
        host.surface()?.setPastePolicy(next);
      },
    },
  };
}
