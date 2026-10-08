// A press on a field the host names selects the field whole, as a press on an equation does,
// and the painted result says so; a field the host does not name keeps the caret at its edge.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import { zipSync, strToU8 } from 'fflate';
import { mountPaginatedSurface, type PaginatedSurface } from '../paginated-surface.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const STYLES_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
// Format defaults, pinned so the fixture does not take the application defaults for omitted docDefaults.
const FORMAT_DOC_DEFAULTS =
  '<w:docDefaults><w:rPrDefault><w:rPr><w:kern w:val="2"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>';

function docx(body: string): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rIdStyles" Type="${STYLES_REL}" Target="styles.xml"/></Relationships>`
    ),
    'word/styles.xml': strToU8(`<w:styles xmlns:w="${W}">${FORMAT_DOC_DEFAULTS}</w:styles>`),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`
    ),
  });
}

/** The page's content box origin, which page-content coordinates are measured from. */
const MARGIN = 72;

interface Mounted {
  readonly surface: PaginatedSurface;
  readonly container: HTMLElement;
  readonly pages: HTMLElement;
}

function mount(body: string, options: { pointer?: 'engine' | 'native' } = {}): Mounted {
  const container = document.createElement('div');
  document.body.append(container);
  const result = mountPaginatedSurface(container, docx(body), {
    scale: 1,
    ...(options.pointer ? { pointer: options.pointer } : {}),
  });
  if (!result.ok) throw new Error(`${result.reason}: ${result.detail ?? ''}`);
  const pages = container.querySelector<HTMLElement>('.docx-pages')!;
  // happy-dom reports no layout at all, so the one measurement the controller makes — where
  // the pages layer sits on screen — is supplied. A deliberately non-zero origin, so a
  // controller that forgot to subtract it would fail rather than pass by coincidence.
  stubRect(pages, { left: 100, top: 50, bottom: 50 });
  return { surface: result.surface, container, pages };
}

function stubRect(
  element: HTMLElement,
  rect: { left: number; top: number; bottom?: number }
): void {
  Object.defineProperty(element, 'getBoundingClientRect', {
    configurable: true,
    value: () => ({
      left: rect.left,
      top: rect.top,
      right: rect.left + 1000,
      bottom: rect.bottom ?? rect.top + 1000,
      width: 1000,
      height: (rect.bottom ?? rect.top + 1000) - rect.top,
      x: rect.left,
      y: rect.top,
    }),
  });
}

/** Page-content coordinates to the client point that lands on them. */
const clientOf = (x: number, y: number) => ({
  clientX: 100 + MARGIN + x,
  clientY: 50 + MARGIN + y,
});

function pointer(type: string, x: number, y: number, init: PointerEventInit = {}): PointerEvent {
  return new PointerEvent(type, {
    bubbles: true,
    cancelable: true,
    button: 0,
    pointerId: 1,
    pointerType: 'mouse',
    ...clientOf(x, y),
    ...init,
  });
}

/** A press, at page-content coordinates. Returns the event so its defaults can be inspected. */
function press(mounted: Mounted, x: number, y: number, init: PointerEventInit = {}): PointerEvent {
  const event = pointer('pointerdown', x, y, init);
  mounted.pages.dispatchEvent(event);
  return event;
}

const release = (x: number, y: number): void => {
  document.dispatchEvent(pointer('pointerup', x, y));
};

const RPR = '<w:rPr><w:sz w:val="22"/></w:rPr>';
const mergeField = (name: string, result: string) =>
  `<w:r>${RPR}<w:fldChar w:fldCharType="begin"/></w:r>` +
  `<w:r>${RPR}<w:instrText xml:space="preserve"> MERGEFIELD "${name}" \\* MERGEFORMAT </w:instrText></w:r>` +
  `<w:r>${RPR}<w:fldChar w:fldCharType="separate"/></w:r>` +
  `<w:r>${RPR}<w:t>${result}</w:t></w:r>` +
  `<w:r>${RPR}<w:fldChar w:fldCharType="end"/></w:r>`;
const text = (value: string) => `<w:r>${RPR}<w:t xml:space="preserve">${value}</w:t></w:r>`;
const BODY = `<w:p>${text('Nome ')}${mergeField('field:a', '«nome»')}${text(' fim')}</w:p>`;

/** The middle of the field's painted result, in page-content coordinates. */
function fieldPoint(surface: PaginatedSurface): [number, number] {
  for (const fragment of surface.layout().pages[0]!.fragments) {
    if (fragment.kind !== 'paragraph') continue;
    for (const line of fragment.lines) {
      const span = line.spans.find((candidate) => candidate.fieldAtom !== undefined);
      if (span) return [span.box.x + span.box.width / 2, span.box.y + span.box.height / 2];
    }
  }
  throw new Error('no field painted');
}

const offsets = (surface: PaginatedSurface): [number, number] => {
  const { anchor, head } = surface.state().selection;
  return [anchor.offset, head.offset];
};

describe('a press on a field', () => {
  test('selects the field whole when the host names its instruction', () => {
    const mounted = mount(BODY);
    mounted.surface.setFieldSelection((instruction) => instruction.includes('"field:'));
    const [x, y] = fieldPoint(mounted.surface);

    press(mounted, x, y);
    release(x, y);

    expect(offsets(mounted.surface)).toEqual([5, 6]);
    const selected = mounted.pages.querySelectorAll('[data-field-atom][data-selected="true"]');
    expect(selected.length).toBeGreaterThan(0);
    mounted.surface.destroy();
  });

  test('keeps the caret at an edge for a field the host does not name', () => {
    const mounted = mount(BODY);
    mounted.surface.setFieldSelection((instruction) => instruction.includes('"other:'));
    const [x, y] = fieldPoint(mounted.surface);

    press(mounted, x, y);
    release(x, y);

    const [anchor, head] = offsets(mounted.surface);
    expect(anchor).toBe(head);
    expect(mounted.pages.querySelectorAll('[data-selected="true"]').length).toBe(0);
    mounted.surface.destroy();
  });

  test('the mark leaves the field when the selection does', () => {
    const mounted = mount(BODY);
    mounted.surface.setFieldSelection(() => true);
    const [x, y] = fieldPoint(mounted.surface);
    press(mounted, x, y);
    release(x, y);

    press(mounted, -40, y);
    release(-40, y);

    expect(offsets(mounted.surface)).toEqual([0, 0]);
    expect(mounted.pages.querySelectorAll('[data-selected="true"]').length).toBe(0);
    mounted.surface.destroy();
  });
});
