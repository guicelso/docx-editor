---
'@docx-editor.dev/pro': patch
---

View-only content-control tags are drawn for block-level controls too: the opening tag at the start of the control's first paragraph and the closing tag at the end of its last, outside the inline tags there. Each tag and caret slot names the control's level (`inline` or `block`), and contentControlEdgesAt takes the part and reports block edges. A key beside a block tag selects the control whole and the next removes it. Removing a block control with its content keeps the invariants deleteBlock keeps: a story keeps a paragraph, a cell ends with one, and a section mark is not dropped.
