// The fields of a content control's boundary record that do not come from geometry: the wrapper's
// own chrome as the file states it, and what the host makes of the control — whether the reader
// meets it as one unit, and the names its chrome publishes.

import type { OoxmlElement } from '@docx-editor.dev/core/store';
import type {
  ContentControlPiece,
  ContentControlPieceDisplay,
  PieceName,
} from '../contracts/editor-content-control-view.ts';
import type { CollectedControl } from './content-control-boundary-layout.ts';
import {
  contentControlPropertiesOf,
  contentControlSubjectOf,
  mapContentControlType,
  propertyChild,
  propertyVal,
} from './content-control-properties.ts';
import { pieceNameRecordOf, type PieceNameRecord } from './piece-names.ts';
import {
  effectiveContentControlLock,
  type ContentControlBoundaryRecord,
} from './semantic-records.ts';

/**
 * Wrapper chrome (alias/tag/type/placeholder/binding), memoized per immutable control node:
 * one settle builds a record per control per pass, and re-reading `w:sdtPr` for hundreds of
 * unchanged controls was a measurable share of the boundary attach.
 */
interface ControlChromeMetadata {
  readonly alias?: string;
  readonly tag?: string;
  readonly controlType: ContentControlBoundaryRecord['controlType'];
  readonly placeholder: boolean;
  readonly bound: boolean;
}

const controlChromeMemos = new WeakMap<OoxmlElement, ControlChromeMetadata>();

function controlChromeOf(control: OoxmlElement): ControlChromeMetadata {
  const cached = controlChromeMemos.get(control);
  if (cached !== undefined) return cached;
  const properties = contentControlPropertiesOf(control);
  const alias = propertyVal(properties, 'alias');
  const tag = propertyVal(properties, 'tag');
  const chrome: ControlChromeMetadata = {
    ...(alias !== undefined ? { alias } : {}),
    ...(tag !== undefined ? { tag } : {}),
    controlType: mapContentControlType(properties),
    placeholder: propertyChild(properties, 'showingPlcHdr') !== undefined,
    bound: propertyChild(properties, 'dataBinding') !== undefined,
  };
  controlChromeMemos.set(control, chrome);
  return chrome;
}

/** A control's record before layout places it: every field but its fragments. */
export function recordWithoutGeometry(
  collected: CollectedControl,
  pieces?: ContentControlPieceDisplay
): ContentControlBoundaryRecord {
  const chrome = controlChromeOf(collected.control);
  return {
    id: collected.control.id,
    ...(chrome.alias !== undefined ? { alias: chrome.alias } : {}),
    ...(chrome.tag !== undefined ? { tag: chrome.tag } : {}),
    controlType: chrome.controlType,
    lock: collected.lockStack[collected.lockStack.length - 1] ?? 'unlocked',
    effectiveLock: effectiveContentControlLock(collected.lockStack),
    placeholder: chrome.placeholder,
    bound: chrome.bound,
    ...hostPieceOf(collected.control, chrome.placeholder, pieces),
    nestingDepth: collected.nestingDepth,
    level: collected.level,
    fragments: [],
  };
}

/** Whether two records carry the same names. */
export function samePieceNames(
  left: PieceNameRecord | undefined,
  right: PieceNameRecord | undefined
): boolean {
  return (
    left?.tone === right?.tone &&
    left?.form === right?.form &&
    left?.state === right?.state &&
    left?.role === right?.role
  );
}

/**
 * The host's answer for a control: an atom or not, and its names. A control showing its
 * placeholder is already a unit — its prompt's — and no atom.
 */
function hostPieceOf(
  control: OoxmlElement,
  showsPlaceholder: boolean,
  display: ContentControlPieceDisplay | undefined
): { readonly atom: boolean; readonly name?: PieceNameRecord } {
  const answer = display?.pieceOf(contentControlSubjectOf(control)) ?? null;
  const name = pieceNameRecordOf(answer === null ? undefined : namesOf(answer));
  const atom = answer !== null && 'atom' in answer && !showsPlaceholder;
  return { atom, ...(name === undefined ? {} : { name }) };
}

function namesOf(piece: ContentControlPiece): PieceName {
  return 'atom' in piece ? piece.atom : piece.frame;
}
