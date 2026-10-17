import type { ExecResult } from './types.ts';

/**
 * How the editor lays the document out: on pages (Word's Print Layout), or in Word's Web Layout
 * (`ST_View` `web`) — one page as wide as the viewport and as tall as the text, with no margins,
 * headers, footers, notes or page borders, where a page or column break opens no page. @public
 */
export type EditorLayoutView = 'print' | 'web';

/** The editor's view, applied without remounting. @public */
export interface EditorViewControl {
  /** The view in force. A requested `web` the document refuses leaves `print` in force. */
  getView(): EditorLayoutView;
  /**
   * Switch views. `web` is refused (`unsupported`, with the reason also in `lastRejection`) for a
   * document with columns, a document grid, notes, floating pictures, floating tables or text
   * frames, and the view in force stays; in `web`, a write that would bring one of those in is
   * refused the same way. The page follows the viewport's content width. The editor keeps the
   * requested view across `load`, `attach` and `detach`, and asks each document it mounts.
   */
  setView(view: EditorLayoutView): ExecResult;
}
