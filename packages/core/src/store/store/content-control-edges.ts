// The inline content-control edges at one offset of a paragraph.
//
// Where controls start or end at the same offset, the offset alone is several places. The edges
// there, in reading order, are what names each place: a caret stands between two of them.

import type { OoxmlNode, OoxmlParagraphNode } from '../package/ooxml-tree.ts';
import { isInlineRunContainer, MAX_INLINE_CONTAINER_DEPTH } from '../package/ooxml-shared.ts';
import { contentControlContentOf, isContentControlNode } from './tree-op-nodes.ts';
import { paragraphOffsetIndex } from './tree-op-segments.ts';

/** One edge of an inline content control. @public */
export interface ContentControlEdge {
  readonly controlId: string;
  readonly edge: 'open' | 'close';
}

/**
 * Every inline control edge at `offset`, in reading order: a control opens before what it holds
 * and closes after it, so `G{B{}E{}}` empty at one offset reads `G open, B open, B close, E open,
 * E close, G close`.
 */
export function contentControlEdgesAt(
  paragraph: OoxmlParagraphNode,
  offset: number
): readonly ContentControlEdge[] {
  const index = paragraphOffsetIndex(paragraph);
  const edges: ContentControlEdge[] = [];
  const walk = (nodes: readonly OoxmlNode[], depth: number): void => {
    if (depth >= MAX_INLINE_CONTAINER_DEPTH) return;
    for (const node of nodes) {
      if (isContentControlNode(node)) {
        const span = index.spanOf(node);
        if (!span || offset < span.start || offset > span.end) continue;
        if (span.start === offset) edges.push({ controlId: node.id, edge: 'open' });
        walk(contentControlContentOf(node)?.children ?? [], depth + 1);
        if (span.end === offset) edges.push({ controlId: node.id, edge: 'close' });
      } else if (isInlineRunContainer(node)) {
        walk(node.children, depth + 1);
      }
    }
  };
  walk(paragraph.children, 0);
  return edges;
}
