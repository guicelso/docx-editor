---
'@docx-editor.dev/pro': patch
---

An edit leaves the caret in the slot where it happened: typed, composed and pasted text keeps the caret beside it, Backspace keeps the edge on the caret's right and Delete the edge on its left, so the next keystroke at an edge where controls meet continues in the same place, with tags drawn or hidden. Composition and paste at a collapsed caret land in its slot, as typing does.
