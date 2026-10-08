---
'@docx-editor.dev/pro': patch
---

A deletion that holds an inline content control and reaches past it takes the control with it, as Word does, instead of leaving an empty zero-width wrapper behind; a wrapper locked against deletion stays, showing its prompt. A deletion that reaches into a control's prompt leaves the prompt whole instead of a few of its letters.
