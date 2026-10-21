import type {
  DocumentChange,
  EditorError,
  EditorEvents,
  EditorHover,
  EditorSnapshot,
  Unsubscribe,
} from '../contracts/editor.ts';
import type { HistoryDiagnostic } from '../contracts/editor-scope.ts';
import type { ResolvedRevisionMarkup } from '../contracts/revision-markup.ts';
import { NO_HOVER, sameHoveredPiece } from './surface-hover.ts';

/** Per-instance notifications. Snapshot derivation stays lazy without subscribers. */
export function createEditorEvents(snapshot: () => EditorSnapshot) {
  const handlers: { [E in keyof EditorEvents]: Set<EditorEvents[E]> } = {
    change: new Set(),
    selectionChange: new Set(),
    error: new Set(),
    historyDiagnostic: new Set(),
    revisionMarkupChange: new Set(),
    hoverChange: new Set(),
  };
  let hover: EditorHover = NO_HOVER;
  const emitHover = (next: EditorHover): void => {
    if (sameHoveredPiece(hover, next)) return;
    hover = next;
    for (const handler of [...handlers.hoverChange]) handler(next);
  };
  return {
    emitError(error: EditorError) {
      for (const handler of [...handlers.error]) handler(error);
    },
    emitDocumentChange(change: DocumentChange) {
      for (const handler of [...handlers.change]) handler(change);
    },
    emitSelectionChange() {
      if (!handlers.selectionChange.size) return;
      const value = snapshot();
      for (const handler of [...handlers.selectionChange]) handler(value);
    },
    emitHistoryDiagnostic(diagnostic: HistoryDiagnostic) {
      for (const handler of [...handlers.historyDiagnostic]) handler(diagnostic);
    },
    emitRevisionMarkupChange(value: ResolvedRevisionMarkup) {
      for (const handler of [...handlers.revisionMarkupChange]) handler(value);
    },
    emitHoverChange: emitHover,
    /** A surface that goes takes its hover with it: no pointer will leave a page that is gone. */
    endHover() {
      emitHover(NO_HOVER);
    },
    hasErrorHandlers: () => handlers.error.size > 0,
    clear() {
      for (const set of Object.values(handlers)) set.clear();
    },
    on<E extends keyof EditorEvents>(event: E, handler: EditorEvents[E]): Unsubscribe {
      handlers[event].add(handler);
      return () => {
        handlers[event].delete(handler);
      };
    },
  };
}
