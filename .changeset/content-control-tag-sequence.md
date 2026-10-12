---
'@docx-editor.dev/pro': patch
---

A content-control tag can be a sequence: `{ variant: 'sequence', labels }` draws several labels one after the other at one edge as ONE tag. It lays out as a single unbreakable unit with the room each label keeps, stands between the edge's two caret slots like any tag, and paints each label as a part of its own with that label's tone and variant. A host whose edge says two things — a condition chip and the text the document prints there — no longer has to choose one: two tags at one edge would put a slot between them that belongs to neither, and let a line break between them. A label with no text is left out, and a sequence left with none draws no tag.
