---
'@docx-editor.dev/pro': patch
---

`session.editOptions()` returns the options the session applies every op with beyond one transaction: the prompt an emptied content control takes, in the reader's language. A host that rehearses ops with `applyTreeOp` before committing them passes it, so the rehearsal writes what the commit will.
