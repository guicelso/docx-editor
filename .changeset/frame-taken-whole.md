---
'@docx-editor.dev/core': minor
---

A host can make a content-control frame a piece the reader takes whole: `{ frame, whole: true }`. A drag or a Shift+click that crosses its edge extends over it, while one inside it stays free; a selection that runs from one edge of it to the other is the selected piece, with the chrome's `data-selected`, the outermost where several share both edges, and the next key removes it. `state().contentControls.selectedControlId` publishes the selected piece.
