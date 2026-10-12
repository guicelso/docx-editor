---
'@docx-editor.dev/pro': patch
---

liftContentControl lifts an inline control that a paragraph holds directly to a block control in its place. The paragraph is cut at the control's edges, through the same split a paragraph mark makes: what comes before stays in it, what comes after goes to a new paragraph with the same properties, and a side with nothing on it is not cut. The control keeps its identity, its properties and every node it holds, now in one paragraph with the properties of the one it came from. Rewriting the control as a new block control would copy what it holds under fresh ids, and an op that follows in the same batch could not address the paragraphs a split mints; this op is how a host changes the level of a control that carries an identity. It refuses a revision, and suggesting mode stamps it so it is refused rather than written outright.
