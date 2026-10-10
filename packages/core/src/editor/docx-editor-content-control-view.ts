/** The host's view over content controls and fields, kept by the facade across surfaces. */

import type {
  ContentControlTagDisplay,
  EditorContentControlView,
  FieldTone,
} from '../contracts/editor-content-control-view.ts';
import type { ContentControlView } from '../layout/content-control-view.ts';
import type { PaginatedSurface, PaginatedSurfaceOptions } from './paginated-surface.ts';

type ViewMountOptions = Pick<
  PaginatedSurfaceOptions,
  'contentControlView' | 'fieldTone' | 'fieldSelection'
>;

/**
 * The view's state and its setters.
 *
 * A surface is rebuilt on `load`, `attach`, `detach` and a refresh recovery. The policies are
 * functions of a control or a field, never of one document, so each rebuilt surface mounts with
 * them, and the live one takes every change.
 */
export function createEditorContentControlView(host: { surface(): PaginatedSurface | null }): {
  readonly mountOptions: () => ViewMountOptions;
  readonly members: EditorContentControlView;
} {
  let view: ContentControlView | null = null;
  let tone: FieldTone | null = null;
  let selection: ((instruction: string) => boolean) | null = null;
  return {
    mountOptions: () => ({
      ...(view ? { contentControlView: view } : {}),
      ...(tone ? { fieldTone: tone } : {}),
      ...(selection ? { fieldSelection: selection } : {}),
    }),
    members: {
      setContentControlTags(display: ContentControlTagDisplay | null) {
        view = display ? { tags: display } : null;
        host.surface()?.setContentControlView(view);
      },
      setFieldTones(next) {
        tone = next;
        host.surface()?.setFieldTones(next);
      },
      setFieldSelection(next) {
        selection = next;
        host.surface()?.setFieldSelection(next);
      },
    },
  };
}
