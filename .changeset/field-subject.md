---
'@docx-editor.dev/core': minor
---

A field is asked by its subject, as a content control is: `FieldSubject = { fieldId, instruction }`, where `fieldId` is the node id of the field's start (its begin `w:fldChar` or its `w:fldSimple`) and names that occurrence across edits around it. `FieldFace` takes the subject, so two fields with the same instruction can take different faces, and `editor.setFieldSelection` takes a `FieldSelectionPolicy` over the subject. The hover's `EditorHoverField` and the hit test's `SemanticHitField` carry the subject too.

`FieldFace` and `setFieldSelection` no longer take the instruction alone.
