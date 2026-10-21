---
'@docx-editor.dev/core': minor
---

A host names the pieces it styles: a tag label, a control and a field carry `tone`, `form` and `state`, published as `data-piece-tone`, `data-piece-form` and `data-piece-state`. `editor.setContentControlPieces` makes a control a named frame or an atom — met as one unit, as a control showing its placeholder is: its plain text lays out as one piece with a chip's room, the caret steps over it, a press selects it whole and the next key removes it. `editor.setFieldFaces` makes a field an atom from its instruction. An atom or a chip may ask for a mark: a slot the engine measures before its text and `editor.setMarkPainter` fills. An atom carries its own `data-hover` and `data-selected`.

`setFieldTones` and `FieldTone` are removed: a field face states the names a tone did, and more. A tag label's `tone` is now one of its names, and `data-cc-tag-tone` and `data-field-tone` are replaced by `data-piece-tone`.
