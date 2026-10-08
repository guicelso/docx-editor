// Where typed text lands beside a prompt.
//
// The store replaces a `w:showingPlcHdr` prompt whole when text is inserted at either of
// its edges, so the characters go where the prompt began, not where the caret was pressed.
// The surface asks the same question the store answers, before the write, so the caret it
// places afterwards counts from the landing rather than from an offset the replacement
// made stale — which had left it past the paragraph's new end, with every later keystroke
// landing in the wrong place.

import type { OoxmlPart, TreeDocOp } from '@docx-editor.dev/core/store';
import type { ContentControlBoundaryRecord } from '../layout/semantic-records.ts';
import {
  promptTypedOver,
  type InlineDestinationFields,
} from '../store/store/tree-op-inline-destination.ts';

type InsertTextOp = Extract<TreeDocOp, { op: 'insertText' }>;

/** The caret offset after `insert` lands: past its text, from where the prompt it replaces began. */
export function promptInsertionLanding(part: OoxmlPart, insert: InsertTextOp): number {
  const prompt = promptTypedOver(part, insert.paragraphId, insert.offset, insert);
  return (prompt?.offset ?? insert.offset) + insert.text.length;
}

/**
 * The control that OWNS a keystroke at the caret, or `undefined` when the text should land
 * where the store's default puts it. A text, date or list control is typed into, so its
 * trailing edge is still "inside". A content-locked or data-bound chip, a checkbox or a
 * picture is an atom: nothing is typed into it, and a caret at either edge types beside it.
 * The store's `refusesTypedContent` names the same kinds and places that text beside the
 * control. It reads each control's own lock, where this reads the inherited one; an inner
 * control under a locked outer one is refused either way.
 */
export function insertOwnerOf(control: ContentControlBoundaryRecord | null): string | undefined {
  if (!control) return undefined;
  if (control.effectiveLock === 'contentLocked' || control.effectiveLock === 'sdtContentLocked') {
    return undefined;
  }
  if (control.bound) return undefined;
  if (control.controlType === 'checkbox' || control.controlType === 'picture') return undefined;
  return control.id;
}

/**
 * The keyboard's `insertText`, landing where `destination` names. Owned by the control the caret
 * sits in, the control's trailing edge means "the end of the field", so typing after the last
 * character of a control stays inside it, as in Word; with no place named, the store's default
 * lands the text beside the control, which is right for a hyperlink and wrong for a form field.
 */
export function typedInsertText(
  target: { readonly paragraphId: string; readonly offset: number },
  text: string,
  destination: InlineDestinationFields
): InsertTextOp {
  return {
    op: 'insertText',
    paragraphId: target.paragraphId,
    offset: target.offset,
    text,
    ...(destination.inside === undefined ? {} : { inside: destination.inside }),
    ...(destination.beside === undefined ? {} : { beside: destination.beside }),
  };
}
