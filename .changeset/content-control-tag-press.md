---
'@docx-editor.dev/pro': patch
---

`editor.setContentControlTagPress` lets a host take a plain press on a view-only content-control tag, found by the layout's hit test and asked in viewing too: a taken press moves no caret, and the page's own listeners still see it.
