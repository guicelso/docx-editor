// The host's view over content controls and fields, as layout projects it, and its cache identity.
//
// The host answers per control and per field. A paragraph's identity is therefore the answers for
// what it draws: the inline controls and the fields inside it, and the block controls that open or
// close at it. It joins the paragraph's projection token, so a changed answer lays out only the
// paragraphs that show it. Folded into the producer instead, one changed label re-measured and
// re-placed every paragraph of the document — the cost the list-count producer once had
// (`semantic-layout.ts`).

import type { OoxmlElement, OoxmlNode, OoxmlPart } from '@docx-editor.dev/core/store';
import type {
  ContentControlPiece,
  ContentControlPieceDisplay,
  ContentControlPromptDisplay,
  ContentControlTagDisplay,
  ContentControlTagLabel,
  FieldFace,
  PieceAtom,
  PieceName,
} from '../contracts/editor-content-control-view.ts';
import { isContentControl } from '../store/package/content-control-walk.ts';
import { parsedFieldSpansOf } from '../store/package/field-nodes.ts';
import { blockControlEdgesOf, type BlockControlEdges } from '../store/store/block-control-edges.ts';
import { contentControlSubjectOf } from './content-control-properties.ts';
import { contentControlPromptOf } from './content-control-prompts.ts';
import { aggregateParagraphTokensForTableBlock, framedTokenJoin } from './layout-cache.ts';
import type { SemanticLayoutOptions } from './semantic-layout-options.ts';

/**
 * What a host shows over content controls and fields, which layout projects and never writes.
 * @public
 *
 * A new object is a new view: layout asks every control and field again and lays out only the
 * paragraphs whose answers changed.
 */
export interface ContentControlView {
  /** The start and end tags at each control. */
  readonly tags?: ContentControlTagDisplay;
  /** The text each control shows while it holds its placeholder. */
  readonly prompts?: ContentControlPromptDisplay;
  /** The piece each control is: an atom, or a frame with the host's names. */
  readonly pieces?: ContentControlPieceDisplay;
  /** The face of each field, which makes it an atom. */
  readonly fieldFaces?: FieldFace;
}

/** What a lane forwards to project the view: the view, and the block edges of its part. */
export interface ContentControlViewFlow {
  readonly contentControlView?: ContentControlView;
  readonly blockControlEdges?: ReadonlyMap<string, BlockControlEdges>;
}

/**
 * The view inputs a lane hands the next one. The two travel together: block edges mean nothing
 * without a view to project at them, and a view without them would draw no block tag.
 */
export function contentControlViewFlow(inputs: ContentControlViewFlow): ContentControlViewFlow {
  if (!inputs.contentControlView) return {};
  return inputs.blockControlEdges
    ? { contentControlView: inputs.contentControlView, blockControlEdges: inputs.blockControlEdges }
    : { contentControlView: inputs.contentControlView };
}

/** The layout options a view adds: its block edges, and its identity in every projection key. */
export type ContentControlViewLayoutOptions = Pick<
  SemanticLayoutOptions,
  | 'blockControlEdges'
  | 'projectionTokenForParagraph'
  | 'projectionTokenForTable'
  | 'projectionEpoch'
>;

/**
 * Join the view's identity onto the projection identities `options` already carries.
 *
 * Done where layout reads the part, so every caller gets it, the document coordinator and a
 * direct `layoutSemanticDocument` call alike. A caller that keys paragraphs without a part
 * epoch keeps the recompute path: the prepass memo must not trust a token it cannot see move.
 */
export function contentControlViewLayoutOptions(
  options: Pick<SemanticLayoutOptions, 'contentControlView'> & ContentControlViewLayoutOptions,
  part: OoxmlPart
): ContentControlViewLayoutOptions {
  const view = options.contentControlView;
  if (!view) return {};
  const edges = blockControlEdgesOf(part);
  const viewToken = (paragraph: OoxmlNode): string => {
    const inline = inlineToken(paragraph, view);
    const block = blockToken(edges.get(paragraph.id), view);
    const fields = fieldsToken(paragraph, view);
    return inline === '' && block === '' && fields === ''
      ? ''
      : `cc-view:${framedTokenJoin([inline, block, fields])}`;
  };
  const own = options.projectionTokenForParagraph;
  const ownTable = options.projectionTokenForTable;
  const tokenForParagraph = (paragraph: OoxmlNode): string =>
    joined(own?.(paragraph) ?? '', viewToken(paragraph));
  const keepsRecomputePath = own !== undefined && options.projectionEpoch === undefined;
  return {
    blockControlEdges: edges,
    projectionTokenForParagraph: tokenForParagraph,
    projectionTokenForTable: (table) =>
      joined(
        ownTable?.(table) ?? (own ? aggregateParagraphTokensForTableBlock(table, own) : ''),
        aggregateParagraphTokensForTableBlock(table, viewToken)
      ),
    ...(keepsRecomputePath
      ? {}
      : {
          projectionEpoch: joined(
            options.projectionEpoch ?? '',
            `cc-view:${versionOf(view)}:${edgesToken(edges)}`
          ),
        }),
  };
}

/** A projection token with the view's joined on, or the token itself when the view adds none. */
function joined(token: string, view: string): string {
  return view === '' ? token : framedTokenJoin([token, view]);
}

/**
 * The answers for the inline controls under one paragraph node, per view.
 *
 * Keyed by the node: inline controls live inside it, so an unchanged node holds the same
 * controls. Every descendant is walked, a hosted text box included, because a host paragraph
 * whose break hits the cache never lays its box out again.
 */
const inlineTokens = new WeakMap<OoxmlNode, WeakMap<ContentControlView, string>>();

function inlineToken(paragraph: OoxmlNode, view: ContentControlView): string {
  let byView = inlineTokens.get(paragraph);
  const cached = byView?.get(view);
  if (cached !== undefined) return cached;
  const parts: string[] = [];
  const stack: OoxmlNode[] = [paragraph];
  while (stack.length > 0) {
    const node = stack.pop()!;
    if (node.kind === 'textValue') continue;
    if (node !== paragraph && isContentControl(node)) parts.push(controlToken(node, view));
    for (let index = node.children.length - 1; index >= 0; index -= 1) {
      stack.push(node.children[index]!);
    }
  }
  const token = parts.length === 0 ? '' : framedTokenJoin(parts);
  if (!byView) inlineTokens.set(paragraph, (byView = new WeakMap()));
  byView.set(view, token);
  return token;
}

/** The answers for the fields of one paragraph node, per view, in document order. */
const fieldTokens = new WeakMap<OoxmlNode, WeakMap<ContentControlView, string>>();

function fieldsToken(paragraph: OoxmlNode, view: ContentControlView): string {
  const face = view.fieldFaces;
  if (!face || paragraph.kind !== 'paragraph') return '';
  let byView = fieldTokens.get(paragraph);
  const cached = byView?.get(view);
  if (cached !== undefined) return cached;
  const parts = parsedFieldSpansOf(paragraph).map((span) => atomToken(face(span.instruction)));
  const token = parts.every((part) => part === '') ? '' : framedTokenJoin(parts);
  if (!byView) fieldTokens.set(paragraph, (byView = new WeakMap()));
  byView.set(view, token);
  return token;
}

/** The answers for the block controls that open and close at one paragraph. */
function blockToken(edges: BlockControlEdges | undefined, view: ContentControlView): string {
  if (!edges) return '';
  return framedTokenJoin([
    framedTokenJoin(edges.opens.map((control) => controlToken(control, view))),
    framedTokenJoin(edges.closes.map((control) => controlToken(control, view))),
  ]);
}

function controlToken(control: OoxmlElement, view: ContentControlView): string {
  const subject = contentControlSubjectOf(control);
  const labels = view.tags?.labelsOf(subject) ?? null;
  return framedTokenJoin([
    control.id,
    labelToken(labels?.open),
    labelToken(labels?.close),
    contentControlPromptOf(control, view.prompts) ?? '',
    pieceToken(view.pieces?.pieceOf(subject) ?? null),
  ]);
}

function labelToken(label: ContentControlTagLabel | undefined): string {
  if (!label) return '';
  if (label.variant === 'edge') return framedTokenJoin(['', '', 'edge']);
  if (label.variant === 'sequence') {
    return framedTokenJoin(['sequence', ...label.labels.map((each) => labelToken(each))]);
  }
  return framedTokenJoin([label.text, nameToken(label), label.mark ?? '', label.variant ?? 'chip']);
}

function pieceToken(piece: ContentControlPiece | null): string {
  if (piece === null) return '';
  return 'atom' in piece
    ? framedTokenJoin(['atom', atomToken(piece.atom)])
    : framedTokenJoin(['frame', nameToken(piece.frame)]);
}

function atomToken(atom: PieceAtom | undefined): string {
  return atom === undefined ? '' : framedTokenJoin([nameToken(atom), atom.mark ?? '']);
}

function nameToken(name: PieceName): string {
  return framedTokenJoin([name.tone ?? '', name.form ?? '', name.state ?? '', name.role ?? '']);
}

/** Which paragraphs every block control opens and closes at, once per edges index. */
const edgesTokens = new WeakMap<ReadonlyMap<string, BlockControlEdges>, string>();

function edgesToken(edges: ReadonlyMap<string, BlockControlEdges>): string {
  const cached = edgesTokens.get(edges);
  if (cached !== undefined) return cached;
  const parts: string[] = [];
  for (const [paragraphId, at] of edges) {
    parts.push(
      framedTokenJoin([
        paragraphId,
        framedTokenJoin(at.opens.map((control) => control.id)),
        framedTokenJoin(at.closes.map((control) => control.id)),
      ])
    );
  }
  const token = framedTokenJoin(parts);
  edgesTokens.set(edges, token);
  return token;
}

/** Each view object's number: a new object is a new view, whatever it holds. */
const viewVersions = new WeakMap<ContentControlView, number>();
let nextViewVersion = 1;

function versionOf(view: ContentControlView): number {
  let version = viewVersions.get(view);
  if (version === undefined) {
    version = nextViewVersion;
    nextViewVersion += 1;
    viewVersions.set(view, version);
  }
  return version;
}
