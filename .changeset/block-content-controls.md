---
'@docx-editor.dev/pro': patch
---

wrapBlocksInContentControl wraps sibling blocks — paragraphs and block-level controls — in a new rich-text block control, and insertBlockContentControl writes one at a place between blocks (before or after a block, or inside a block control that shows its prompt), holding the paragraphs given or a paragraph showing the prompt. removeContentControl unwraps either. A tracked block control has no implementation yet, so both ops refuse a revision, and suggesting mode stamps them so they are refused rather than written outright.
