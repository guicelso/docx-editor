/** The host's policies the facade keeps across surfaces: over controls, edits and writes. */

import { createEditorContentControlView } from './docx-editor-content-control-view.ts';
import { createEditorEditPolicy } from './docx-editor-edit-policy.ts';
import { createEditorWriteAdmission } from './docx-editor-write-admission.ts';
import type { PaginatedSurface } from './paginated-surface.ts';

/** Every policy's mount options for a new surface, and every setter the editor publishes. */
export function createEditorHostPolicies(host: { surface(): PaginatedSurface | null }) {
  const view = createEditorContentControlView(host);
  const edits = createEditorEditPolicy(host);
  const writes = createEditorWriteAdmission(host);
  return {
    mountOptions: () => ({
      ...view.mountOptions(),
      ...edits.mountOptions(),
      ...writes.mountOptions(),
    }),
    members: { ...view.members, ...edits.members, ...writes.members },
  };
}
