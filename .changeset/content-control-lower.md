---
'@docx-editor.dev/pro': patch
---

lowerContentControl lowers a block control holding exactly one paragraph into that paragraph, the inverse of liftContentControl. The paragraph takes the control's place among the blocks, keeping its identity and properties, and holds the control, which keeps its identity, its properties and every node the paragraph held, its prompt included. Joining the paragraph with its neighbours stays with joinParagraphs, by the ids the lowering keeps, so a host can bring a control and the controls it holds back into one line within one batch. A control holding anything but one paragraph, or one already in a line, is unsupported; it refuses a revision, and suggesting mode stamps it so it is refused rather than written outright.
