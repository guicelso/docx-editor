---
'@docx-editor.dev/core': minor
---

A view-only content-control tag no longer glues to the word beside it. A line can wrap before an open tag and after a close tag, as around an inline object; the tags of nested controls stay together, and so does punctuation that cannot end or start a line. Tags that travel together share the room the line leaves them: a tag wider than their common share keeps the start of its longest label and ends in an ellipsis. Line breaking, the anchor-line probe and table autofit share the same seam rule, so a chain of tags with no space between them no longer stands past the margin.
