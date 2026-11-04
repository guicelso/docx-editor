/** The host's policies the facade keeps across surfaces: over controls, edits, writes and pastes. */

import { createEditorContentControlView } from './docx-editor-content-control-view.ts';
import { createEditorEditPolicy } from './docx-editor-edit-policy.ts';
import { createEditorPastePolicy } from './docx-editor-paste-policy.ts';
import { createEditorWriteAdmission } from './docx-editor-write-admission.ts';
import type { PaginatedSurface } from './paginated-surface.ts';

/** Every policy's mount options for a new surface, and every setter the editor publishes. */
export function createEditorHostPolicies(host: { surface(): PaginatedSurface | null }) {
  const view = createEditorContentControlView(host);
  const edits = createEditorEditPolicy(host);
  const writes = createEditorWriteAdmission(host);
  const pastes = createEditorPastePolicy(host);
  return {
    mountOptions: () => ({
      ...view.mountOptions(),
      ...edits.mountOptions(),
      ...writes.mountOptions(),
      ...pastes.mountOptions(),
    }),
    members: { ...view.members, ...edits.members, ...writes.members, ...pastes.members },
  };
}
