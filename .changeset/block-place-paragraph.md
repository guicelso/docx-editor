---
'@docx-editor.dev/pro': patch
---

insertParagraph writes a paragraph before or after a block — a paragraph, a table or a block-level control — holding the text given, with the paragraph properties of the edge it stands beside and the face of the run there; the caret goes to it. The caret slot outside a block control's tag is that place: typing there opens a paragraph with the text, and Enter opens an empty one. Pasting there is refused rather than landed inside the control.
