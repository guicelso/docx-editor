// Where each BLOCK-level content control opens and closes: at the start of its first paragraph and
// the end of its last, in document order.
//
// Layout flattens block controls into the flow, so a paragraph's own walk cannot tell that it is the
// first or last of one. This index answers it for a whole part, once per part revision: parts are
// immutable and every edit publishes a new root, so the memo can never answer for another revision.
// It is keyed by the PART, never by a paragraph node — the same node is shared by the revisions an
// edit leaves untouched, and an index of one would answer for the other.

import {
  contentControlContentNodeOf,
  contentControlLevelOf,
  MAX_CONTENT_CONTROL_NESTING,
} from '../package/content-control-nodes.ts';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';

/** The block controls that open and close at one paragraph. @public */
export interface BlockControlEdges {
  /** Controls whose first paragraph this is, outermost first. */
  readonly opens: readonly OoxmlElement[];
  /** Controls whose last paragraph this is, innermost first. */
  readonly closes: readonly OoxmlElement[];
}

const indexes = new WeakMap<OoxmlNode, ReadonlyMap<string, BlockControlEdges>>();

/** Every paragraph a block control opens or closes at, by paragraph id. @public */
export function blockControlEdgesOf(part: OoxmlPart): ReadonlyMap<string, BlockControlEdges> {
  const cached = indexes.get(part.root);
  if (cached) return cached;
  const index = collectBlockControlEdges(part.root);
  indexes.set(part.root, index);
  return index;
}

interface MutableEdges {
  readonly opens: OoxmlElement[];
  readonly closes: OoxmlElement[];
}

function collectBlockControlEdges(root: OoxmlNode): ReadonlyMap<string, BlockControlEdges> {
  const index = new Map<string, MutableEdges>();
  const at = (paragraphId: string): MutableEdges => {
    let edges = index.get(paragraphId);
    if (!edges) {
      edges = { opens: [], closes: [] };
      index.set(paragraphId, edges);
    }
    return edges;
  };
  let pending: OoxmlElement[] = [];
  let lastParagraph: string | null = null;
  let paragraphs = 0;

  const walk = (nodes: readonly OoxmlNode[], nesting: number): void => {
    for (const node of nodes) {
      if (node.kind === 'textValue') continue;
      if (node.kind === 'paragraph') {
        if (pending.length > 0) at(node.id).opens.push(...pending);
        pending = [];
        lastParagraph = node.id;
        paragraphs += 1;
        continue;
      }
      if (node.kind === 'contentControl' && contentControlLevelOf(node) === 'block') {
        if (nesting >= MAX_CONTENT_CONTROL_NESTING) continue;
        visitControl(node, nesting);
        continue;
      }
      walk(node.children, nesting);
    }
  };

  const visitControl = (control: OoxmlElement, nesting: number): void => {
    const before = paragraphs;
    pending.push(control);
    walk(contentControlContentNodeOf(control)?.children ?? [], nesting + 1);
    if (paragraphs > before && lastParagraph !== null) {
      at(lastParagraph).closes.push(control);
      return;
    }
    pending = pending.filter((open) => open !== control);
  };

  walk([root], 0);
  return index;
}
