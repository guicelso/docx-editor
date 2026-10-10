---
'@docx-editor.dev/pro': patch
---

A split at the edge of a complex field keeps the field whole. The pieces of a field after its `begin` (the instruction, the separator, the result and the `end`) measured nothing to the split, so a paragraph split, a many-way split, an inline control inserted right after a field, or a control divided there left the `begin` on one side and the rest of the field on the other. Every run holding a piece of the field now goes with it.
