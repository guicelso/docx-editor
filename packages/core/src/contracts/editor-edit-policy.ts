/**
 * The edits a host may take over before the engine writes them: what a paragraph break or a
 * deletion means where the host's own structure stands.
 */

/**
 * A structural edit the user asked for, by key or by input method alike. @public
 *
 * `paragraphBreak` is Enter, and an input method's paragraph insertion. `delete` is Backspace
 * (`backward`) or Delete (`forward`), by character or, with the word modifier or the input
 * method's word deletion, by word. A line break, a page break and typed text are not asked.
 */
export type EditIntent =
  | { readonly kind: 'paragraphBreak' }
  | {
      readonly kind: 'delete';
      readonly direction: 'backward' | 'forward';
      readonly unit: 'character' | 'word';
    };

/**
 * The host's answer to an edit intent. `handled` means the host wrote what the intent means
 * here, or decided nothing is written, and the engine writes nothing; `default` lets the engine
 * do what it always does. It runs on the edit's path, before the engine resolves it, with the
 * selection the edit would act on already in the surface's state. @public
 */
export type EditPolicy = (intent: EditIntent) => 'handled' | 'default';

/** The host's policy over structural edits. @public */
export interface EditorEditPolicy {
  /**
   * Ask `policy` before every paragraph break and deletion; `null` removes it. The editor keeps
   * the policy across `load`, `attach` and `detach`, and hands it to each document it mounts.
   */
  setEditPolicy(policy: EditPolicy | null): void;
}
