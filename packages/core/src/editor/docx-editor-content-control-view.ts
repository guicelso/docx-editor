/** The host's view over content controls and fields, kept by the facade across surfaces. */

import type {
  ContentControlPieceDisplay,
  ContentControlPromptDisplay,
  ContentControlTagDisplay,
  ContentControlTagPressPolicy,
  EditorContentControlView,
  FieldFace,
  FieldSelectionPolicy,
} from '../contracts/editor-content-control-view.ts';
import type { MarkPainter } from '../output/semantic-paint-pieces.ts';
import type { DocxEditorInstance } from './docx-editor-types.ts';
import type { ContentControlView } from '../layout/content-control-view.ts';
import type { PaginatedSurface, PaginatedSurfaceOptions } from './paginated-surface.ts';

type ViewMountOptions = Pick<
  PaginatedSurfaceOptions,
  'contentControlView' | 'markPainter' | 'fieldSelection' | 'contentControlTagPress'
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
  readonly members: EditorContentControlView & Pick<DocxEditorInstance, 'setMarkPainter'>;
} {
  let tags: ContentControlTagDisplay | null = null;
  let prompts: ContentControlPromptDisplay | null = null;
  let pieces: ContentControlPieceDisplay | null = null;
  let fieldFaces: FieldFace | null = null;
  let view: ContentControlView | null = null;
  let markPainter: MarkPainter | null = null;
  let selection: FieldSelectionPolicy | null = null;
  let tagPress: ContentControlTagPressPolicy | null = null;
  // Each change is a new view object: layout asks every control and field again and keeps what is
  // unchanged.
  const install = (): void => {
    const answers: ContentControlView = {
      ...(tags ? { tags } : {}),
      ...(prompts ? { prompts } : {}),
      ...(pieces ? { pieces } : {}),
      ...(fieldFaces ? { fieldFaces } : {}),
    };
    view = Object.keys(answers).length > 0 ? answers : null;
    host.surface()?.setContentControlView(view);
  };
  return {
    mountOptions: () => ({
      ...(view ? { contentControlView: view } : {}),
      ...(markPainter ? { markPainter } : {}),
      ...(selection ? { fieldSelection: selection } : {}),
      ...(tagPress ? { contentControlTagPress: tagPress } : {}),
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
      setContentControlPieces(display) {
        pieces = display;
        install();
      },
      setFieldFaces(face) {
        fieldFaces = face;
        install();
      },
      setMarkPainter(next) {
        markPainter = next;
        host.surface()?.setMarkPainter(next);
      },
      setFieldSelection(next) {
        selection = next;
        host.surface()?.setFieldSelection(next);
      },
      setContentControlTagPress(next) {
        tagPress = next;
        host.surface()?.setContentControlTagPress(next);
      },
    },
  };
}
