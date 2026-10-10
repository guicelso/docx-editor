/** The host's policy over structural edits, kept by the facade across surfaces. */

import type { EditPolicy, EditorEditPolicy } from '../contracts/editor-edit-policy.ts';
import type { PaginatedSurface, PaginatedSurfaceOptions } from './paginated-surface.ts';

/**
 * The policy and its setter. A surface is rebuilt on `load`, `attach`, `detach` and a refresh
 * recovery; the policy answers for the host's structure, never for one document, so each rebuilt
 * surface mounts with it, and the live one takes every change.
 */
export function createEditorEditPolicy(host: { surface(): PaginatedSurface | null }): {
  readonly mountOptions: () => Pick<PaginatedSurfaceOptions, 'editPolicy'>;
  readonly members: EditorEditPolicy;
} {
  let policy: EditPolicy | null = null;
  return {
    mountOptions: () => (policy ? { editPolicy: policy } : {}),
    members: {
      setEditPolicy(next) {
        policy = next;
        host.surface()?.setEditPolicy(next);
      },
    },
  };
}
