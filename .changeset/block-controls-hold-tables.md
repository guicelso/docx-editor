---
'@docx-editor.dev/pro': patch
---

A block-level content control holds tables. `wrapBlocksInContentControl` wraps a run of sibling blocks that includes tables, and `insertBlockContentControl` takes `blocks` — paragraphs, tables and block-level controls — where it took `paragraphs`. A control that starts or ends with a table opens at its first cell paragraph and closes at its last, the paragraphs its tags stand on in reading order; its frame already covered the whole table.
