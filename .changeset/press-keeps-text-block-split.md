---
'@docx-editor.dev/core': minor
---

A press on text where content-control edges meet keeps the caret with the text it pressed: at the end of a control's last word the caret is inside the control, at the start of the word after it, outside. The hit test publishes the side as `SemanticHit.pressedText`. `splitContentControl` also divides a block control `at` a place in a paragraph it holds (`BlockControlPlace`), strictly inside its content: the start of a paragraph moves it to the tail, its end moves the blocks after it, and a place inside cuts it as `splitParagraph` does. The head keeps the control, and the tail is a new control with the tag named and a fresh `w:id`; the blocks move with their ids.
