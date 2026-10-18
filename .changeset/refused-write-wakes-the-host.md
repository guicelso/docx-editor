---
'@docx-editor.dev/pro': patch
---

A refused write now emits `selectionChange` when its reason moves `snapshot().lastRejection`, so a host learns that the keystroke, paste or command it just saw did nothing, and a toolbar can say why.
