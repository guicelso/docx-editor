---
'@docx-editor.dev/pro': patch
---

insertMergeField authors a `MERGEFIELD "<name>"` at a caret — five runs carrying the caret's formatting, landing in the paragraph or in the place `inside` or `beside` names — and setMergeField rewrites a merge field's name and cached result in place, keeping its runs and their formatting. mergeFieldsOf reads a paragraph's merge fields with their names and offsets.
