// Word's Web Layout view (`ST_View` `web`, ECMA-376 Part 1 `wml.xsd`): one sheet as wide as
// its host and as tall as its content. Word confirms no page breaks, headers or footers there,
// and says nothing about columns, the line grid, notes or content positioned against the
// page, so a story holding any of them is refused rather than approximated.

import type { OoxmlNode, OoxmlPart } from '@docx-editor.dev/core/store';
import { isFloatingLegacyVmlAtom } from '../store/package/legacy-vml-projection.ts';
import { isNoteReferenceNode } from '../store/package/note-nodes.ts';
import { WP_NAMESPACE_URI } from '../store/package/ooxml-shared.ts';
import type { PageGeometry } from './page-geometry.ts';
import { hasParagraphFrames } from './paragraph-frame-exclusion.ts';
import { enumerateDocumentSectionsFromBlocks } from './section-properties.ts';
import { storyBlocks } from './story-roots.ts';
import type { StyleCascadeTable } from './style-cascade.ts';
import { hasFloatingTables } from './table-float-exclusion.ts';

/** How a story is laid out: on sheets, or as one sheet `widthPt` wide with no bottom. */
export type LayoutView =
  | { readonly kind: 'print' }
  | { readonly kind: 'web'; readonly widthPt: number };

/** What the web view does not lay out, by the first one a story holds. */
export type WebLayoutViewRefusal =
  | 'columns'
  | 'line-grid'
  | 'notes'
  | 'floating-drawing'
  | 'floating-table'
  | 'text-frame';

/** The sheet of the web view: no margins and no bottom, so no line ever leaves it. */
export function webLayoutGeometry(widthPt: number): PageGeometry {
  return Object.freeze({
    width: widthPt,
    height: Number.POSITIVE_INFINITY,
    margin: Object.freeze({ top: 0, right: 0, bottom: 0, left: 0 }),
    headerDistance: 0,
    footerDistance: 0,
  });
}

const inlineRefusals = new WeakMap<OoxmlNode, WebLayoutViewRefusal | null>();

function isAnchoredDrawing(node: OoxmlNode): boolean {
  if (node.kind === 'anchoredDrawing') return true;
  return (
    node.kind === 'generic' && node.localName === 'anchor' && node.namespaceUri === WP_NAMESPACE_URI
  );
}

function ownInlineRefusal(node: OoxmlNode): WebLayoutViewRefusal | null {
  if (isNoteReferenceNode(node)) return 'notes';
  if (isAnchoredDrawing(node) || isFloatingLegacyVmlAtom(node)) return 'floating-drawing';
  return null;
}

/**
 * The first note reference or floating drawing under `node`, memoized by node identity: an
 * edit replaces only the path to what it changed, so a commit rescans that path alone.
 */
function inlineRefusalIn(node: OoxmlNode): WebLayoutViewRefusal | null {
  if (node.kind === 'textValue') return null;
  const memo = inlineRefusals.get(node);
  if (memo !== undefined) return memo;
  let refusal = ownInlineRefusal(node);
  for (const child of node.children) {
    if (refusal !== null) break;
    refusal = inlineRefusalIn(child);
  }
  inlineRefusals.set(node, refusal);
  return refusal;
}

/**
 * Why the web view cannot lay `part` out, or `null` when it can.
 *
 * Frames and floating tables are read through the same predicates the paged layout uses, so a
 * style that floats a table is refused like a direct `w:tblpPr`.
 */
export function webLayoutViewRefusal(
  part: OoxmlPart,
  styles: StyleCascadeTable | undefined,
  widthPt: number
): WebLayoutViewRefusal | null {
  const blocks = storyBlocks(part, 'all-markup', undefined, styles);
  for (const section of enumerateDocumentSectionsFromBlocks(part, blocks).sections) {
    if (section.properties.columns.count > 1) return 'columns';
    if (section.properties.gridLinePitchTwips !== undefined) return 'line-grid';
  }
  if (hasParagraphFrames(blocks, styles)) return 'text-frame';
  if (hasFloatingTables(blocks, widthPt, styles, 'all-markup', undefined)) return 'floating-table';
  return inlineRefusalIn(part.root);
}
