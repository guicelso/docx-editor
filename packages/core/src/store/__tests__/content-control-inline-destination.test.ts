// A new control or an inline fragment lands where the caller names: beside a control, or inside one.
//
// One model offset at a tagged edge is several places, and an offset strictly inside an inline
// control has no paragraph-level sibling position at all. `inside` and `beside` are the
// destinations `insertText` already takes; these pin that a control and a fragment honour them.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, type OoxmlNode, type OoxmlPart } from '../package/ooxml-tree.ts';
import { applyTreeOp } from '../store/tree-op-apply.ts';
import { validateTreeOp } from '../store/tree-op-validate.ts';
import type { TreeDocOp } from '../store/tree-op-types.ts';
import { contentControlPropertiesOf } from '../package/content-control-nodes.ts';
import {
  isContentControl,
  contentControlContentChildren,
} from '../package/content-control-walk.ts';
import { storyBlocks } from '../../layout/story-roots.ts';
import { contentControlSubjectOf } from '../../layout/content-control-properties.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const sdt = (tag: string, inner: string, properties = '') =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/>${properties}<w:richText/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

/** `CPF G{B{RG}E{CNH}} fim`: offsets 4 (G and B open), 6 (B closes, E opens), 9 (E and G close). */
function documentOf(
  body = `${run('CPF ')}${sdt('G', sdt('B', run('RG')) + sdt('E', run('CNH')))}${run(' fim')}`
) {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body><w:p>${body}</w:p></w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'app/xml' }
  );
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

/** The paragraph with its controls bracketed by tag; a control with no tag is `P`. */
function bracketed(part: OoxmlPart): string {
  const paragraph = storyBlocks(part)[0]!;
  const walk = (nodes: readonly OoxmlNode[]): string =>
    nodes
      .map((node) => {
        if (node.kind === 'textValue') return node.value;
        if (isContentControl(node)) {
          return `${contentControlSubjectOf(node).tag ?? 'P'}{${walk(contentControlContentChildren(node))}}`;
        }
        return walk(node.children);
      })
      .join('');
  return walk(paragraph.kind === 'textValue' ? [] : paragraph.children);
}

function controlNamed(part: OoxmlPart, tag: string): OoxmlNode {
  let found: OoxmlNode | null = null;
  const walk = (nodes: readonly OoxmlNode[]): void => {
    for (const node of nodes) {
      if (node.kind === 'textValue') continue;
      if (isContentControl(node) && contentControlSubjectOf(node).tag === tag) found = node;
      walk(node.children);
    }
  };
  walk([part.root]);
  if (!found) throw new Error(`no control ${tag}`);
  return found;
}

type Destination =
  | { readonly inside: string }
  | { readonly beside: { readonly tag: string; readonly side: 'before' | 'after' } }
  | null;

function destinationOf(part: OoxmlPart, destination: Destination) {
  if (destination === null) return {};
  return 'inside' in destination
    ? { inside: controlNamed(part, destination.inside).id }
    : {
        beside: {
          controlId: controlNamed(part, destination.beside.tag).id,
          side: destination.beside.side,
        },
      };
}

const PROMPT = (type: string) => `<${type}>`;

/** Apply, and confirm that validation answers what the applier did. */
function applied(part: OoxmlPart, op: TreeDocOp): OoxmlPart {
  const rejection = validateTreeOp(part, op);
  const result = applyTreeOp(part, op, { placeholderPrompt: PROMPT });
  if (!result.ok) {
    expect(rejection).toBe(result.reason);
    throw new Error(`refused: ${result.reason}`);
  }
  expect(rejection).toBeNull();
  return result.part;
}

function controlAt(
  start: number,
  end: number,
  destination: Destination,
  part = documentOf()
): string {
  const op: TreeDocOp = {
    op: 'insertContentControl',
    paragraphId: storyBlocks(part)[0]!.id,
    start,
    end,
    type: 'richText',
    ...destinationOf(part, destination),
  };
  return bracketed(applied(part, op));
}

function fragmentOf(inner: string): OoxmlNode {
  const read = readOoxmlPart(`<w:p xmlns:w="${W}">${inner}</w:p>`, {
    name: '/fragment.xml',
    contentType: 'application/xml',
  });
  if (!read.ok) throw new Error(read.reason);
  return read.part.root;
}

function fragmentAt(
  offset: number,
  destination: Destination,
  blocks = [fragmentOf(run('X'))]
): string {
  const part = documentOf();
  const op: TreeDocOp = {
    op: 'insertFragment',
    paragraphId: storyBlocks(part)[0]!.id,
    offset,
    blocks,
    ...destinationOf(part, destination),
  };
  return bracketed(applied(part, op));
}

describe('an empty control at a caret, by destination', () => {
  test('with no destination the paragraph rule holds: an edge goes outside, inside a control is refused', () => {
    expect(controlAt(4, 4, null)).toBe('CPF P{<richText>}G{B{RG}E{CNH}} fim');
    expect(() => controlAt(5, 5, null)).toThrow('refused: indivisible-content');
  });

  test('beside G, before: in front of the group', () => {
    expect(controlAt(4, 4, { beside: { tag: 'G', side: 'before' } })).toBe(
      'CPF P{<richText>}G{B{RG}E{CNH}} fim'
    );
  });

  test('beside B, before: in the group, before the first branch', () => {
    expect(controlAt(4, 4, { beside: { tag: 'B', side: 'before' } })).toBe(
      'CPF G{P{<richText>}B{RG}E{CNH}} fim'
    );
  });

  test('beside B, after: in the group, between the branches', () => {
    expect(controlAt(6, 6, { beside: { tag: 'B', side: 'after' } })).toBe(
      'CPF G{B{RG}P{<richText>}E{CNH}} fim'
    );
  });

  test('beside G, after: behind the group', () => {
    expect(controlAt(9, 9, { beside: { tag: 'G', side: 'after' } })).toBe(
      'CPF G{B{RG}E{CNH}}P{<richText>} fim'
    );
  });

  test('inside B at its start, in its middle and at its end', () => {
    expect(controlAt(4, 4, { inside: 'B' })).toBe('CPF G{B{P{<richText>}RG}E{CNH}} fim');
    expect(controlAt(5, 5, { inside: 'B' })).toBe('CPF G{B{RP{<richText>}G}E{CNH}} fim');
    expect(controlAt(6, 6, { inside: 'B' })).toBe('CPF G{B{RGP{<richText>}}E{CNH}} fim');
  });

  test('inside G at the shared edge: between the branches, not in either', () => {
    expect(controlAt(6, 6, { inside: 'G' })).toBe('CPF G{B{RG}P{<richText>}E{CNH}} fim');
  });

  test('the new control shows the prompt the transaction names, under w:showingPlcHdr', () => {
    const part = documentOf();
    const after = applied(part, {
      op: 'insertContentControl',
      paragraphId: storyBlocks(part)[0]!.id,
      start: 5,
      end: 5,
      type: 'richText',
      tag: 'N',
      inside: controlNamed(part, 'B').id,
    });
    expect(contentControlPropertiesOf(controlNamed(after, 'N')).showingPlaceholder).toBe(true);
    expect(bracketed(after)).toBe('CPF G{B{RN{<richText>}G}E{CNH}} fim');
  });

  test('inside a control showing its prompt, the new control replaces the prompt', () => {
    const empty = applied(documentOf(), {
      op: 'insertContentControl',
      paragraphId: storyBlocks(documentOf())[0]!.id,
      start: 4,
      end: 4,
      type: 'richText',
      tag: 'O',
      beside: { controlId: controlNamed(documentOf(), 'G').id, side: 'before' },
    });
    const paragraphId = storyBlocks(empty)[0]!.id;
    const nested = applied(empty, {
      op: 'insertContentControl',
      paragraphId,
      start: 7,
      end: 7,
      type: 'richText',
      tag: 'N',
      inside: controlNamed(empty, 'O').id,
    });
    expect(bracketed(nested)).toBe('CPF O{N{<richText>}}G{B{RG}E{CNH}} fim');
    expect(contentControlPropertiesOf(controlNamed(nested, 'O')).showingPlaceholder).toBe(false);
    expect(contentControlPropertiesOf(controlNamed(nested, 'N')).showingPlaceholder).toBe(true);
    // No empty run of the prompt stays beside the new control.
    expect(
      contentControlContentChildren(controlNamed(nested, 'O')).map((node) => node.kind)
    ).toEqual(['contentControl']);
  });

  test('a stale destination is refused, not landed elsewhere', () => {
    expect(() => controlAt(6, 6, { beside: { tag: 'G', side: 'before' } })).toThrow(
      'refused: unknown-content-control'
    );
    expect(() => controlAt(2, 2, { inside: 'B' })).toThrow('refused: offset-out-of-range');
  });

  test('a content-locked owner refuses the control', () => {
    const locked = documentOf(
      `${run('a')}${sdt('L', run('bc'), '<w:lock w:val="contentLocked"/>')}`
    );
    expect(() => controlAt(2, 2, { inside: 'L' }, locked)).toThrow('refused: locked');
  });
});

describe('a range wrapped inside a control', () => {
  test('the characters of a branch', () => {
    expect(controlAt(4, 6, { inside: 'B' })).toBe('CPF G{B{P{RG}}E{CNH}} fim');
    expect(controlAt(7, 8, { inside: 'E' })).toBe('CPF G{B{RG}E{CP{N}H}} fim');
  });

  test('the branches of a group, whole', () => {
    expect(controlAt(4, 9, { inside: 'G' })).toBe('CPF G{P{B{RG}E{CNH}}} fim');
  });

  test('a range that crosses the owner edge, or cuts a nested control, is refused', () => {
    expect(() => controlAt(3, 6, { inside: 'B' })).toThrow('refused: offset-out-of-range');
    expect(() => controlAt(5, 8, { inside: 'G' })).toThrow('refused: indivisible-content');
  });

  test('beside names a caret, never a range', () => {
    expect(() => controlAt(4, 6, { beside: { tag: 'B', side: 'before' } })).toThrow(
      'refused: invalidArgs'
    );
  });
});

describe('an inline fragment, by destination', () => {
  test('beside G, after, and inside B in its middle', () => {
    expect(fragmentAt(9, { beside: { tag: 'G', side: 'after' } })).toBe('CPF G{B{RG}E{CNH}}X fim');
    expect(fragmentAt(5, { inside: 'B' })).toBe('CPF G{B{RXG}E{CNH}} fim');
    expect(fragmentAt(6, { beside: { tag: 'B', side: 'after' } })).toBe('CPF G{B{RG}XE{CNH}} fim');
  });

  test('a fragment holding a control lands it whole in the destination', () => {
    expect(fragmentAt(6, { inside: 'G' }, [fragmentOf(sdt('F', run('Y')))])).toBe(
      'CPF G{B{RG}F{Y}E{CNH}} fim'
    );
  });

  test('a fragment of more than one paragraph has no inline destination', () => {
    expect(() =>
      fragmentAt(5, { inside: 'B' }, [fragmentOf(run('X')), fragmentOf(run('Y'))])
    ).toThrow('refused: invalidArgs');
  });
});
