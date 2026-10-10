---
'@docx-editor.dev/pro': patch
---

splitContentControl cuts an inline control in two at an offset strictly inside its content: the head keeps the control and the tail is a new control with the tag the caller names and a fresh `w:id`. joinContentControls puts the content of the second of two adjacent sibling inline controls at the end of the first and removes the second; a control showing its prompt holds nothing to keep. A paragraph split copies the tag of a control it crosses, which leaves two controls claiming one identity; these ops are how a host divides or joins a control that carries one. Both refuse a revision, and suggesting mode stamps them so they are refused rather than written outright.
