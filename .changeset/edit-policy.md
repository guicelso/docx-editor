---
'@docx-editor.dev/pro': patch
---

`editor.setEditPolicy(policy)` lets a host take over a paragraph break or a deletion before the engine writes it. The policy is asked with the intent — `{ kind: 'paragraphBreak' }`, or `{ kind: 'delete', direction, unit }` — for Enter, Backspace and Delete and for an input method's paragraph insertion and deletions alike, so a host's structure is honoured whichever way the edit arrives. `handled` leaves the document to the host; `default` lets the engine write what it always writes. A line break, a page break and typed text are not asked. The editor keeps the policy across `load`, `attach` and `detach`.
