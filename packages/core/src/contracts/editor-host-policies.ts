import type { EditorContentControlView } from './editor-content-control-view.ts';
import type { EditorEditPolicy } from './editor-edit-policy.ts';
import type { EditorPastePolicy } from './editor-paste-policy.ts';
import type { EditorWriteAdmission } from './editor-write-admission.ts';

export type * from './editor-content-control-view.ts';
export type * from './editor-edit-policy.ts';
export type * from './editor-paste-policy.ts';
export type * from './editor-write-admission.ts';

/**
 * The host's policies over its documents: its view of content controls, what a structural
 * edit means, what a write may publish, and what a paste lands. The editor keeps each across
 * documents. @public
 */
export interface EditorHostPolicies
  extends EditorContentControlView, EditorEditPolicy, EditorWriteAdmission, EditorPastePolicy {}
