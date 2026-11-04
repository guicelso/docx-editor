// The host's say over a rich paste, asked inside the paste's own transaction.

import type { OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';

/** The blocks a paste brings, merged for the document, and the story part as it stands. @public */
export interface PastedBlocks {
  readonly blocks: readonly OoxmlNode[];
  readonly part: OoxmlPart;
}

/** The story part before a paste and as the paste would leave it. @public */
export interface PasteLanding {
  readonly before: OoxmlPart;
  readonly after: OoxmlPart;
}

/**
 * The host's say over a paste of the engine's own fragment or of external HTML. Plain text is not
 * asked. @public
 *
 * `blocksOf` answers the blocks that land in place of the ones brought, for data of the host's
 * own that a copy must not share, such as an identity; the engine's budgets and shape checks
 * still apply to what it answers. `admit` sees the part before and after: a string refuses the
 * paste with that reason, and no lower-fidelity flavour retries it.
 */
export interface PastePolicy {
  readonly blocksOf?: (pasted: PastedBlocks) => readonly OoxmlNode[];
  readonly admit?: (landing: PasteLanding) => string | null;
}
