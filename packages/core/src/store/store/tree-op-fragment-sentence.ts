// A clipboard fragment landing in a sentence (rich-clipboard-fidelity 2.1).
//
// Inside an inline control a paragraph mark has no place, so a fragment that lands there is read
// as one line: each paragraph gives its inline content, a block control keeps its properties and
// takes its paragraphs' content inline, and one space joins what two blocks held — the reading
// the fragment's plain text gets when its lines are joined. A table has no such reading.

import { WML_NAMESPACE_URI, XML_NAMESPACE_URI } from '../package/ooxml-shared.ts';
import type { OoxmlElement, OoxmlNode, OoxmlParagraphNode } from '../package/ooxml-tree.ts';
import { siteBesideControl } from './tree-op-beside.ts';
import { contentControlContentOf, paragraphPropertiesNodeOf } from './tree-op-nodes.ts';
import type { InlineDestinationFields } from './tree-op-inline-destination.ts';

/**
 * Whether a named place is in a sentence: inside an inline control, or beside a control whose
 * parent is not the paragraph itself. Beside a control that stands in the paragraph, blocks still
 * have a place: the paragraph divides there.
 */
export function placeIsSentence(
  paragraph: OoxmlParagraphNode,
  offset: number,
  fields: InlineDestinationFields
): boolean {
  if (fields.inside !== undefined) return true;
  if (fields.beside === undefined) return false;
  const site = siteBesideControl(paragraph, offset, fields.beside);
  return site?.kind === 'newRun' && site.holder.id !== paragraph.id;
}

/** Whether every block has a sentence reading: nothing that holds a table. */
export function takesSentence(blocks: readonly OoxmlNode[]): boolean {
  return blocks.every((block) => !holdsTable(block));
}

/** The fragment as one line of inline content; ids for the joining spaces come from `nextId`. */
export function sentenceOf(blocks: readonly OoxmlNode[], nextId: () => string): OoxmlNode[] {
  return joined(
    blocks.map((block) => inlineOf(block, nextId)),
    nextId
  );
}

/** A paragraph's inline children, keeping the bindings the paragraph declared for them. */
export function inlineChildrenOf(paragraph: OoxmlElement): readonly OoxmlNode[] {
  const pPr = paragraphPropertiesNodeOf(paragraph);
  return paragraph.children
    .filter((child) => child !== pPr)
    .map((child) => withParentBindings(child, paragraph));
}

/** Keep local bindings when a paste removes the node's original parent. */
export function withParentBindings(node: OoxmlNode, parent: OoxmlElement): OoxmlNode {
  if (node.kind === 'textValue' || parent.namespaceBindings.length === 0) return node;
  const ownPrefixes = new Set(node.namespaceBindings.map((binding) => binding.prefix));
  const inherited = parent.namespaceBindings.filter((binding) => !ownPrefixes.has(binding.prefix));
  if (inherited.length === 0) return node;
  return { ...node, namespaceBindings: [...inherited, ...node.namespaceBindings] };
}

function holdsTable(node: OoxmlNode): boolean {
  if (node.kind === 'textValue') return false;
  return node.kind === 'table' || node.children.some(holdsTable);
}

function inlineOf(block: OoxmlNode, nextId: () => string): readonly OoxmlNode[] {
  if (block.kind === 'textValue') return [];
  if (block.kind === 'paragraph') return inlineChildrenOf(block);
  if (block.kind !== 'contentControl') return [block];
  const content = contentControlContentOf(block);
  if (!content) return [block];
  const inline = joined(
    content.children.map((child) => inlineOf(child, nextId)),
    nextId
  );
  return [
    {
      ...block,
      children: block.children.map((child) =>
        child.id === content.id ? { ...content, children: inline } : child
      ),
    } as OoxmlNode,
  ];
}

function joined(parts: readonly (readonly OoxmlNode[])[], nextId: () => string): OoxmlNode[] {
  return parts
    .filter((part) => part.length > 0)
    .flatMap((part, index) => (index === 0 ? [...part] : [spaceRun(nextId), ...part]));
}

function spaceRun(nextId: () => string): OoxmlNode {
  return {
    id: nextId(),
    kind: 'run',
    namespaceUri: WML_NAMESPACE_URI,
    localName: 'r',
    prefix: 'w',
    namespaceBindings: [],
    attributes: [],
    children: [
      {
        id: nextId(),
        kind: 'text',
        namespaceUri: WML_NAMESPACE_URI,
        localName: 't',
        prefix: 'w',
        namespaceBindings: [],
        attributes: [
          {
            kind: 'xmlSpace',
            namespaceUri: XML_NAMESPACE_URI,
            localName: 'space',
            prefix: 'xml',
            value: 'preserve',
          },
        ],
        children: [{ id: nextId(), kind: 'textValue', value: ' ' }],
      },
    ],
  } as unknown as OoxmlNode;
}
