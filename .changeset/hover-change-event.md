---
'@docx-editor.dev/pro': patch
---

Publish the piece under a resting pointer on its own `hoverChange` event: the innermost content control, the field the pointer is on, and where that piece is painted. A control keeps the hover across its value widget. The hover no longer rides `selectionChange`, and `state().contentControls.hoveredControlId` is removed: it re-derived the snapshot for every selection subscriber on each pointer move. One layout hit test now drives both the event and the `data-hover` chrome; the second, DOM-based hover tracker is removed.
