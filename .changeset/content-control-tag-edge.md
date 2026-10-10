---
'@docx-editor.dev/pro': patch
---

A content-control tag can be an edge: `{ variant: 'edge' }` draws nothing and takes no room, and still stands as a tag, so the place just inside the control is a caret slot of its own on that side. A host that draws a label on one side of a control and nothing on the other no longer loses the inside of the undrawn side: in a list whose items draw no label where they open, typing right after a separator lands in the next item rather than between the two. `ContentControlTagLabel` is now a union of the drawn label (`ContentControlTagInk`) and the edge (`ContentControlTagEdgeOnly`).
