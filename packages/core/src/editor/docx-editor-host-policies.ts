/** The host's policies the facade keeps across surfaces: its view over controls, and over edits. */

import { createEditorContentControlView } from './docx-editor-content-control-view.ts';
import { createEditorEditPolicy } from './docx-editor-edit-policy.ts';
import type { PaginatedSurface } from './paginated-surface.ts';

/** Every policy's mount options for a new surface, and every setter the editor publishes. */
export function createEditorHostPolicies(host: { surface(): PaginatedSurface | null }) {
  const view = createEditorContentControlView(host);
  const edits = createEditorEditPolicy(host);
  return {
    mountOptions: () => ({ ...view.mountOptions(), ...edits.mountOptions() }),
    members: { ...view.members, ...edits.members },
  };
}
