// The host's view over content controls, as layout projects it, and its cache identity.
//
// The host answers per control. A paragraph's identity is therefore the answers for the controls
// it draws: the inline controls inside it, and the block controls that open or close at it. It
// joins the paragraph's projection token, so a changed answer lays out only the paragraphs that
// show it. Folded into the producer instead, one changed label re-measured and re-placed every
// paragraph of the document — the cost the list-count producer once had (`semantic-layout.ts`).

import type { OoxmlElement, OoxmlNode, OoxmlPart } from '@docx-editor.dev/core/store';
import type {
  ContentControlPromptDisplay,
  ContentControlTagDisplay,
  ContentControlTagLabel,
} from '../contracts/editor-content-control-view.ts';
import { isContentControl } from '../store/package/content-control-walk.ts';
import { blockControlEdgesOf, type BlockControlEdges } from '../store/store/block-control-edges.ts';
import { contentControlSubjectOf } from './content-control-properties.ts';
import { contentControlPromptOf } from './content-control-prompts.ts';
import { aggregateParagraphTokensForTableBlock, framedTokenJoin } from './layout-cache.ts';
import type { SemanticLayoutOptions } from './semantic-layout-options.ts';

/**
 * What a host shows over content controls, which layout projects and never writes. @public
 *
 * A new object is a new view: layout asks every control again and lays out only the paragraphs
 * whose answers changed.
 */
export interface ContentControlView {
  /** The start and end tags at each control. */
  readonly tags?: ContentControlTagDisplay;
  /** The text each control shows while it holds its placeholder. */
  readonly prompts?: ContentControlPromptDisplay;
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
    return inline === '' && block === '' ? '' : `cc-view:${framedTokenJoin([inline, block])}`;
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

/** The answers for the block controls that open and close at one paragraph. */
function blockToken(edges: BlockControlEdges | undefined, view: ContentControlView): string {
  if (!edges) return '';
  return framedTokenJoin([
    framedTokenJoin(edges.opens.map((control) => controlToken(control, view))),
    framedTokenJoin(edges.closes.map((control) => controlToken(control, view))),
  ]);
}

function controlToken(control: OoxmlElement, view: ContentControlView): string {
  const labels = view.tags?.labelsOf(contentControlSubjectOf(control)) ?? null;
  return framedTokenJoin([
    control.id,
    labelToken(labels?.open),
    labelToken(labels?.close),
    contentControlPromptOf(control, view.prompts) ?? '',
  ]);
}

function labelToken(label: ContentControlTagLabel | undefined): string {
  return label ? framedTokenJoin([label.text, label.tone ?? '']) : '';
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
