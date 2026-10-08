---
'@docx-editor.dev/pro': patch
---

Text inserted at a named place replaces only that place's prompt: `inside` a control replaces the control's own prompt, `beside` a control replaces none, so the prompt of a neighbour sharing the offset is no longer typed over and a write in front of a control is no longer refused.
