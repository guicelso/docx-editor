// A control is cut in two — an inline one at an offset, a block one at a place in a paragraph — or two inline
// ones are joined, without a character changing: the head keeps the control's identity, the tail is
// a new control with the tag the caller names, and a join puts the second's content after the first's.

import { describe, expect, test } from 'bun:test';
import { contentControlPropertiesOf } from '../package/content-control-nodes.ts';
import { readOoxmlPart, type OoxmlNode, type OoxmlPart } from '../package/ooxml-tree.ts';
import { applyTreeOp } from '../store/tree-op-apply.ts';
import { treeOpReach } from '../store/tree-op-content-controls.ts';
import type { TreeDocOp } from '../store/tree-op-types.ts';
import { validateTreeOp } from '../store/tree-op-validate.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const inline = (tag: string, inner: string, extra = '') =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/>${extra}<w:richText/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
const prompt = (tag: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w:showingPlcHdr/><w:richText/></w:sdtPr><w:sdtContent>${run('type here')}</w:sdtContent></w:sdt>`;
const field =
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> MERGEFIELD x </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>val</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';

function documentOf(paragraph: string): OoxmlPart {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body><w:p>${paragraph}</w:p><w:sectPr/></w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'app/xml' }
  );
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

/** The paragraph as `tag[content]` for every control, text as it reads. */
function reading(part: OoxmlPart): string {
  const walk = (nodes: readonly OoxmlNode[]): string =>
    nodes
      .map((node) => {
        if (node.kind === 'textValue') return node.value;
        if (node.kind === 'contentControl') {
          const content = node.children.find((child) => child.kind === 'contentControlContent');
          const inner = content && content.kind !== 'textValue' ? content.children : [];
          return `${contentControlPropertiesOf(node).tag}[${walk(inner)}]`;
        }
        if (node.kind === 'contentControlProperties' || node.localName === 'instrText') return '';
        return walk(node.children);
      })
      .join('');
  return walk(paragraphOf(part).children);
}

function paragraphOf(part: OoxmlPart): Exclude<OoxmlNode, { kind: 'textValue' }> {
  const body = part.root.kind === 'textValue' ? null : part.root.children[0];
  const paragraph = body && body.kind !== 'textValue' ? body.children[0] : undefined;
  if (!paragraph || paragraph.kind === 'textValue') throw new Error('no paragraph');
  return paragraph;
}

function controlNamed(part: OoxmlPart, tag: string): Exclude<OoxmlNode, { kind: 'textValue' }> {
  let found: Exclude<OoxmlNode, { kind: 'textValue' }> | null = null;
  const walk = (node: OoxmlNode): void => {
    if (found || node.kind === 'textValue') return;
    if (node.kind === 'contentControl' && contentControlPropertiesOf(node).tag === tag)
      found = node;
    else node.children.forEach(walk);
  };
  walk(part.root);
  if (found === null) throw new Error(`no control ${tag}`);
  return found;
}

const idOf = (part: OoxmlPart, tag: string) => controlNamed(part, tag).id;

function applied(part: OoxmlPart, op: TreeDocOp): OoxmlPart {
  expect(validateTreeOp(part, op)).toBeNull();
  const result = applyTreeOp(part, op);
  if (!result.ok) throw new Error(`refused: ${result.reason}`);
  return result.part;
}

function refused(part: OoxmlPart, op: TreeDocOp): string {
  const rejection = validateTreeOp(part, op);
  const result = applyTreeOp(part, op);
  expect(result.ok).toBe(false);
  expect(rejection).toBe(result.ok ? null : result.reason);
  return rejection ?? '';
}

describe('splitContentControl', () => {
  test('the head keeps the control; the tail is a new control with the tag named', () => {
    const part = documentOf(
      `${run('x ')}${inline('item:a', run('o RGa CNH'), '<w:id w:val="7"/>')}`
    );
    const head = idOf(part, 'item:a');

    const split = applied(part, {
      op: 'splitContentControl',
      controlId: head,
      offset: 6,
      tag: 'item:b',
    });

    expect(reading(split)).toBe('x item:a[o RG]item:b[a CNH]');
    expect(idOf(split, 'item:a')).toBe(head);
    expect(contentControlPropertiesOf(controlNamed(split, 'item:a')).id).toBe(7);
    expect(contentControlPropertiesOf(controlNamed(split, 'item:b')).id).toBe(8);
  });

  test('a control nested in another is divided inside its holder', () => {
    const part = documentOf(inline('list', inline('item:a', run('abcd'))));

    const split = applied(part, {
      op: 'splitContentControl',
      controlId: idOf(part, 'item:a'),
      offset: 2,
      tag: 'item:b',
    });

    expect(reading(split)).toBe('list[item:a[ab]item:b[cd]]');
  });

  test('a control nested in the one divided goes whole to its side', () => {
    const part = documentOf(
      inline('item:a', `${run('ab')}${inline('opt', run('cd'))}${run('ef')}`)
    );

    const split = applied(part, {
      op: 'splitContentControl',
      controlId: idOf(part, 'item:a'),
      offset: 4,
      tag: 'item:b',
    });

    expect(reading(split)).toBe('item:a[abopt[cd]]item:b[ef]');
  });

  test('inside a nested control the offset is not a place', () => {
    const nested = documentOf(inline('item:a', `${run('ab')}${inline('opt', run('cd'))}`));

    expect(
      refused(nested, {
        op: 'splitContentControl',
        controlId: idOf(nested, 'item:a'),
        offset: 3,
        tag: 'b',
      })
    ).toBe('indivisible-content');
  });

  test('a field is one place, and goes whole to its side', () => {
    const part = documentOf(inline('item:a', `${run('ab')}${field}${run('cd')}`));

    const split = applied(part, {
      op: 'splitContentControl',
      controlId: idOf(part, 'item:a'),
      offset: 3,
      tag: 'item:b',
    });

    expect(reading(split)).toBe('item:a[abval]item:b[cd]');
  });

  test('an edge, a prompt, a block control and a revision are refused', () => {
    const part = documentOf(`${inline('item:a', run('abcd'))}${prompt('item:p')}`);
    const controlId = idOf(part, 'item:a');
    const block = readOoxmlPart(
      `<w:document xmlns:w="${W}"><w:body>${inline('blk', `<w:p>${run('abcd')}</w:p>`)}<w:sectPr/></w:body></w:document>`,
      { name: '/word/document.xml', contentType: 'app/xml' }
    );
    if (!block.ok) throw new Error(block.reason);

    expect(refused(part, { op: 'splitContentControl', controlId, offset: 0, tag: 'b' })).toBe(
      'invalid-range'
    );
    expect(refused(part, { op: 'splitContentControl', controlId, offset: 4, tag: 'b' })).toBe(
      'invalid-range'
    );
    expect(
      refused(part, {
        op: 'splitContentControl',
        controlId: idOf(part, 'item:p'),
        offset: 6,
        tag: 'b',
      })
    ).toBe('invalidArgs');
    expect(
      refused(block.part, {
        op: 'splitContentControl',
        controlId: idOf(block.part, 'blk'),
        offset: 2,
        tag: 'b',
      })
    ).toBe('unsupported');
    expect(
      refused(part, {
        op: 'splitContentControl',
        controlId,
        offset: 2,
        tag: 'b',
        revision: { author: 'QA', date: '2026-10-10T00:00:00Z' },
      })
    ).toBe('invalidArgs');
  });

  test('a locked or bound control is not divided', () => {
    const wrapper = documentOf(inline('item:a', run('abcd'), '<w:lock w:val="sdtLocked"/>'));
    const content = documentOf(inline('item:a', run('abcd'), '<w:lock w:val="contentLocked"/>'));
    const bound = documentOf(
      inline(
        'item:a',
        run('abcd'),
        '<w:dataBinding w:xpath="/a" w:storeItemID="{00000000-0000-0000-0000-000000000000}"/>'
      )
    );

    for (const part of [wrapper, content]) {
      expect(
        refused(part, {
          op: 'splitContentControl',
          controlId: idOf(part, 'item:a'),
          offset: 2,
          tag: 'b',
        })
      ).toBe('locked');
    }
    expect(
      refused(bound, {
        op: 'splitContentControl',
        controlId: idOf(bound, 'item:a'),
        offset: 2,
        tag: 'b',
      })
    ).toBe('bound');
  });

  test('a lock on the control holding it reaches the split', () => {
    const part = documentOf(
      inline('list', inline('item:a', run('abcd')), '<w:lock w:val="contentLocked"/>')
    );
    const op: TreeDocOp = {
      op: 'splitContentControl',
      controlId: idOf(part, 'item:a'),
      offset: 2,
      tag: 'b',
    };

    expect(treeOpReach(op)).toEqual({
      kind: 'nodes',
      targets: [{ nodeId: idOf(part, 'item:a'), structural: true }],
    });
    expect(refused(part, op)).toBe('locked');
  });
});

describe('splitContentControl · a block control', () => {
  const p = (text: string) => `<w:p>${run(text)}</w:p>`;
  const listOf = (items: string) =>
    readOoxmlPart(
      `<w:document xmlns:w="${W}"><w:body>${inline('list', items)}<w:sectPr/></w:body></w:document>`,
      { name: '/word/document.xml', contentType: 'app/xml' }
    );
  const bodyOf = (part: OoxmlPart): string => {
    const walk = (nodes: readonly OoxmlNode[]): string =>
      nodes
        .map((node) => {
          if (node.kind === 'textValue') return node.value;
          if (node.kind === 'contentControlProperties') return '';
          if (node.kind === 'contentControl') {
            const content = node.children.find((child) => child.kind === 'contentControlContent');
            return `${contentControlPropertiesOf(node).tag}[${walk(content && content.kind !== 'textValue' ? content.children : [])}]`;
          }
          return node.kind === 'paragraph' ? `(${walk(node.children)})` : walk(node.children);
        })
        .join('');
    return walk(part.root.kind === 'textValue' ? [] : part.root.children);
  };
  const blockIdOf = (part: OoxmlPart, tag: string, index: number): string => {
    const content = controlNamed(part, tag).children.find(
      (child) => child.kind === 'contentControlContent'
    );
    const block = content && content.kind !== 'textValue' ? content.children[index] : undefined;
    if (!block) throw new Error(`no block ${index} in ${tag}`);
    return block.id;
  };

  const at = (part: OoxmlPart, index: number, offset: number) => ({
    paragraphId: blockIdOf(part, 'item:a', index),
    offset,
  });
  const split = (
    items: string,
    place: (part: OoxmlPart) => { paragraphId: string; offset: number }
  ) => {
    const read = listOf(inline('item:a', items, '<w:id w:val="7"/>'));
    if (!read.ok) throw new Error(read.reason);
    const head = idOf(read.part, 'item:a');
    const written = applied(read.part, {
      op: 'splitContentControl',
      controlId: head,
      at: place(read.part),
      tag: 'item:b',
    });
    return { read: read.part, written, head };
  };

  test('at the start of a paragraph, the paragraph goes whole to a new control; the head keeps the control', () => {
    const { read, written, head } = split(`${p('B1')}${p('B2')}${p('B3')}`, (part) =>
      at(part, 1, 0)
    );

    expect(bodyOf(written)).toBe('list[item:a[(B1)]item:b[(B2)(B3)]]');
    expect(idOf(written, 'item:a')).toBe(head);
    expect(blockIdOf(written, 'item:b', 0)).toBe(blockIdOf(read, 'item:a', 1));
    expect(contentControlPropertiesOf(controlNamed(written, 'item:a')).id).toBe(7);
    expect(contentControlPropertiesOf(controlNamed(written, 'item:b')).id).toBe(8);
  });

  test('at the end of a paragraph, the blocks after it go, and it stays whole', () => {
    const { written } = split(`${p('B1')}${p('B2')}`, (part) => at(part, 0, 2));

    expect(bodyOf(written)).toBe('list[item:a[(B1)]item:b[(B2)]]');
  });

  test('inside a paragraph, the paragraph is cut, and its tail opens the new control', () => {
    const { written } = split(`${p('que fica livre')}${p('B2')}`, (part) => at(part, 0, 4));

    expect(bodyOf(written)).toBe('list[item:a[(que )]item:b[(fica livre)(B2)]]');
  });

  test('the very start and end, a paragraph it does not hold, a nested control, a prompt, an inline control and a revision are refused', () => {
    const read = listOf(
      `${inline('item:a', `${p('B1')}<w:p>${inline('nested', run('xyz'))}</w:p>`)}${prompt('item:p')}`
    );
    if (!read.ok) throw new Error(read.reason);
    const controlId = idOf(read.part, 'item:a');
    const inlinePart = documentOf(inline('item:i', run('abcd')));
    const refusal = (place: { paragraphId: string; offset: number }, id = controlId) =>
      refused(read.part, { op: 'splitContentControl', controlId: id, at: place, tag: 'b' });

    expect(refusal(at(read.part, 0, 0))).toBe('invalid-range');
    expect(refusal(at(read.part, 1, 3))).toBe('invalid-range');
    expect(refusal({ paragraphId: idOf(read.part, 'item:p'), offset: 0 })).toBe('invalid-range');
    expect(refusal(at(read.part, 1, 1))).toBe('indivisible-content');
    expect(
      refusal(
        { paragraphId: blockIdOf(read.part, 'item:a', 0), offset: 1 },
        idOf(read.part, 'item:p')
      )
    ).toBe('invalidArgs');
    expect(
      refused(read.part, {
        op: 'splitContentControl',
        controlId,
        at: at(read.part, 0, 1),
        tag: 'b',
        revision: { author: 'QA', date: '2026-10-10T00:00:00Z' },
      })
    ).toBe('invalidArgs');
    expect(
      refused(inlinePart, {
        op: 'splitContentControl',
        controlId: idOf(inlinePart, 'item:i'),
        at: { paragraphId: idOf(inlinePart, 'item:i'), offset: 1 },
        tag: 'b',
      })
    ).toBe('unsupported');
  });
});

describe('joinContentControls', () => {
  test("the second control's content goes to the end of the first, and the second goes", () => {
    const part = documentOf(
      inline('list', `${inline('item:a', run('o RG'))}${inline('item:b', run('a CNH'))}`)
    );

    const joined = applied(part, {
      op: 'joinContentControls',
      firstId: idOf(part, 'item:a'),
      secondId: idOf(part, 'item:b'),
    });

    expect(reading(joined)).toBe('list[item:a[o RGa CNH]]');
  });

  test('a prompt holds nothing to keep, on either side', () => {
    const emptySecond = documentOf(`${inline('item:a', run('ab'))}${prompt('item:b')}`);
    const emptyFirst = documentOf(`${prompt('item:a')}${inline('item:b', run('cd'))}`);

    const keptFirst = applied(emptySecond, {
      op: 'joinContentControls',
      firstId: idOf(emptySecond, 'item:a'),
      secondId: idOf(emptySecond, 'item:b'),
    });
    const filledFirst = applied(emptyFirst, {
      op: 'joinContentControls',
      firstId: idOf(emptyFirst, 'item:a'),
      secondId: idOf(emptyFirst, 'item:b'),
    });

    expect(reading(keptFirst)).toBe('item:a[ab]');
    expect(reading(filledFirst)).toBe('item:a[cd]');
    expect(contentControlPropertiesOf(controlNamed(filledFirst, 'item:a')).showingPlaceholder).toBe(
      false
    );
  });

  test('two prompts join into the first prompt', () => {
    const part = documentOf(`${prompt('item:a')}${prompt('item:b')}`);

    const joined = applied(part, {
      op: 'joinContentControls',
      firstId: idOf(part, 'item:a'),
      secondId: idOf(part, 'item:b'),
    });

    expect(reading(joined)).toBe('item:a[type here]');
    expect(contentControlPropertiesOf(controlNamed(joined, 'item:a')).showingPlaceholder).toBe(
      true
    );
  });

  test('controls that are not adjacent siblings are not joined', () => {
    const apart = documentOf(
      `${inline('item:a', run('ab'))}${run(', ')}${inline('item:b', run('cd'))}`
    );
    const reversed = documentOf(`${inline('item:a', run('ab'))}${inline('item:b', run('cd'))}`);

    expect(
      refused(apart, {
        op: 'joinContentControls',
        firstId: idOf(apart, 'item:a'),
        secondId: idOf(apart, 'item:b'),
      })
    ).toBe('not-adjacent-siblings');
    expect(
      refused(reversed, {
        op: 'joinContentControls',
        firstId: idOf(reversed, 'item:b'),
        secondId: idOf(reversed, 'item:a'),
      })
    ).toBe('not-adjacent-siblings');
  });

  test('a second control locked against removal is not joined', () => {
    const part = documentOf(
      `${inline('item:a', run('ab'))}${inline('item:b', run('cd'), '<w:lock w:val="sdtLocked"/>')}`
    );

    expect(
      refused(part, {
        op: 'joinContentControls',
        firstId: idOf(part, 'item:a'),
        secondId: idOf(part, 'item:b'),
      })
    ).toBe('locked');
  });
});
