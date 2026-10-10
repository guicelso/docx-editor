/** The host's view over content controls and fields, kept by the facade across surfaces. */

import type {
  ContentControlPromptDisplay,
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
  let tags: ContentControlTagDisplay | null = null;
  let prompts: ContentControlPromptDisplay | null = null;
  let view: ContentControlView | null = null;
  let tone: FieldTone | null = null;
  let selection: ((instruction: string) => boolean) | null = null;
  // Each change is a new view object: layout asks every control again and keeps what is unchanged.
  const install = (): void => {
    view = tags || prompts ? { ...(tags ? { tags } : {}), ...(prompts ? { prompts } : {}) } : null;
    host.surface()?.setContentControlView(view);
  };
  return {
    mountOptions: () => ({
      ...(view ? { contentControlView: view } : {}),
      ...(tone ? { fieldTone: tone } : {}),
      ...(selection ? { fieldSelection: selection } : {}),
    }),
    members: {
      setContentControlTags(display) {
        tags = display;
        install();
      },
      setContentControlPrompts(display) {
        prompts = display;
        install();
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
