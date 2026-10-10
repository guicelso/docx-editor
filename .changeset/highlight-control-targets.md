---
'@docx-editor.dev/core': minor
---

setHighlights takes content controls as well as text ranges: a `{ controlId }` target marks the control whole — its content in every paragraph it reaches and the tags at its edges and its children's, the area its boundary outlines — at either level. A control mark stops painting when the document no longer holds the control, and getHighlightsAt reports it with `controlId` in place of `start` and `length`. HighlightHit keeps its shape for range targets.
