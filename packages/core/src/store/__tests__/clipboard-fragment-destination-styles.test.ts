// A fragment merged with the destination's styles, the way Word's "Use Destination Styles" lands a
// building block: a style the target has by name is the target's, one it lacks is imported, and the
// fragment's defaults are not stamped onto the content.
import { describe, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '../package/ooxml-tree.ts';
import { mergeFragmentIntoPackage } from '../store/clipboard-fragment-merge.ts';
import type { OoxmlNode } from '../package/ooxml-tree.ts';
import { W, buildPackage } from './clipboard-fragment-fixtures.ts';

const styles = (inner: string): string => `<w:styles xmlns:w="${W}">${inner}</w:styles>`;
const heading = (id: string, color: string): string =>
  `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="heading 1"/>` +
  `<w:rPr><w:color w:val="${color}"/></w:rPr></w:style>`;
const tenPoint =
  '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="20"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>';
const twelvePoint =
  '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>';

function merged(
  targetStyles: string,
  fragmentStyles: string,
  fragmentBody: string,
  styleSource: 'source' | 'destination'
): { readonly blocks: string; readonly targetStyles: string } {
  const target = buildPackage('<w:p/>', { 'word/styles.xml': targetStyles });
  const fragment = buildPackage(fragmentBody, { 'word/styles.xml': fragmentStyles });
  const result = mergeFragmentIntoPackage(target, fragment, target.mainDocumentPart, {
    styles: styleSource,
  });
  if (!result.ok) throw new Error(result.reason);
  const host = result.pkg.parts.get(result.pkg.mainDocumentPart)!;
  const blocks = serializeOoxmlPart({
    ...host,
    root: { ...host.root, children: result.blocks as OoxmlNode[] } as typeof host.root,
  });
  return {
    blocks,
    targetStyles: serializeOoxmlPart(result.pkg.parts.get('/word/styles.xml')!),
  };
}

const titled =
  '<w:p><w:pPr><w:pStyle w:val="Ttulo1"/></w:pPr><w:r><w:t>Do pedido</w:t></w:r></w:p>';

describe('the destination styles policy', () => {
  test('a style the target has by name is the target one, whatever the id', () => {
    const result = merged(
      styles(tenPoint + heading('Heading1', '0000FF')),
      styles(twelvePoint + heading('Ttulo1', 'FF0000')),
      titled,
      'destination'
    );
    expect(result.blocks).toContain('<w:pStyle w:val="Heading1"/>');
    expect(result.targetStyles).not.toContain('FF0000');
    expect(result.targetStyles).not.toContain('Ttulo1');
  });

  test('a style the target lacks is imported under its own id', () => {
    const result = merged(
      styles(tenPoint),
      styles(twelvePoint + heading('Ttulo1', 'FF0000')),
      titled,
      'destination'
    );
    expect(result.blocks).toContain('<w:pStyle w:val="Ttulo1"/>');
    expect(result.targetStyles).toContain('w:styleId="Ttulo1"');
  });

  test('a taken id under another name is imported under a fresh id', () => {
    const other =
      '<w:style w:type="paragraph" w:styleId="Ttulo1"><w:name w:val="caption"/></w:style>';
    const result = merged(
      styles(tenPoint + other),
      styles(twelvePoint + heading('Ttulo1', 'FF0000')),
      titled,
      'destination'
    );
    expect(result.blocks).not.toContain('<w:pStyle w:val="Ttulo1"/>');
    expect(result.blocks).toContain('<w:pStyle w:val="Ttulo1Pasted"/>');
    expect(result.targetStyles).toContain('FF0000');
  });

  test('the same name of another type is not the same style', () => {
    const character =
      '<w:style w:type="character" w:styleId="Heading1Char"><w:name w:val="heading 1"/></w:style>';
    const result = merged(
      styles(tenPoint + character),
      styles(twelvePoint + heading('Ttulo1', 'FF0000')),
      titled,
      'destination'
    );
    expect(result.blocks).toContain('<w:pStyle w:val="Ttulo1"/>');
    expect(result.targetStyles).toContain('w:styleId="Ttulo1"');
  });

  test('the fragment defaults are not stamped onto the content', () => {
    const plain = '<w:p><w:r><w:t>Do pedido</w:t></w:r></w:p>';
    const destination = merged(styles(tenPoint), styles(twelvePoint), plain, 'destination');
    const source = merged(styles(tenPoint), styles(twelvePoint), plain, 'source');
    expect(destination.blocks).not.toContain('<w:sz');
    expect(source.blocks).toContain('<w:sz w:val="24"/>');
  });

  test('the source policy keeps the fragment look, as a paste does', () => {
    const result = merged(
      styles(tenPoint + heading('Ttulo1', '0000FF')),
      styles(twelvePoint + heading('Ttulo1', 'FF0000')),
      titled,
      'source'
    );
    expect(result.blocks).toContain('<w:pStyle w:val="Ttulo1Pasted"/>');
    expect(result.targetStyles).toContain('FF0000');
  });
});
