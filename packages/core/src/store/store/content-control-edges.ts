// The content-control edges at one offset of a paragraph.
//
// Where controls start or end at the same offset, the offset alone is several places. The edges
// there, in reading order, are what names each place: a caret stands between two of them. A block
// control opens at the start of its first paragraph and closes at the end of its last, outside
// every inline edge there.

import type { OoxmlNode, OoxmlParagraphNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { blockControlEdgesOf } from './block-control-edges.ts';
import { isInlineRunContainer, MAX_INLINE_CONTAINER_DEPTH } from '../package/ooxml-shared.ts';
import { contentControlContentOf, isContentControlNode } from './tree-op-nodes.ts';
import { paragraphOffsetIndex } from './tree-op-segments.ts';

/** One edge of a content control. @public */
export interface ContentControlEdge {
  readonly controlId: string;
  readonly edge: 'open' | 'close';
  /** Inside the paragraph, or holding whole paragraphs. */
  readonly level: 'inline' | 'block';
}

/**
 * Every control edge at `offset`, in reading order: a control opens before what it holds and closes
 * after it, so `G{B{}E{}}` empty at one offset reads `G open, B open, B close, E open, E close,
 * G close`. The block controls the paragraph opens come first, at its start; the ones it closes come
 * last, at its end.
 */
export function contentControlEdgesAt(
  part: OoxmlPart,
  paragraph: OoxmlParagraphNode,
  offset: number
): readonly ContentControlEdge[] {
  const index = paragraphOffsetIndex(paragraph);
  const block = blockControlEdgesOf(part).get(paragraph.id);
  const edges: ContentControlEdge[] = [];
  if (block && offset === 0) {
    for (const control of block.opens)
      edges.push({ controlId: control.id, edge: 'open', level: 'block' });
  }
  const walk = (nodes: readonly OoxmlNode[], depth: number): void => {
    if (depth >= MAX_INLINE_CONTAINER_DEPTH) return;
    for (const node of nodes) {
      if (isContentControlNode(node)) {
        const span = index.spanOf(node);
        if (!span || offset < span.start || offset > span.end) continue;
        if (span.start === offset)
          edges.push({ controlId: node.id, edge: 'open', level: 'inline' });
        walk(contentControlContentOf(node)?.children ?? [], depth + 1);
        if (span.end === offset) edges.push({ controlId: node.id, edge: 'close', level: 'inline' });
      } else if (isInlineRunContainer(node)) {
        walk(node.children, depth + 1);
      }
    }
  };
  walk(paragraph.children, 0);
  if (block && offset === index.length) {
    for (const control of block.closes)
      edges.push({ controlId: control.id, edge: 'close', level: 'block' });
  }
  return edges;
}
