---
'@docx-editor.dev/core': minor
---

A paste lands where the caret names, as typing does: the slot where control edges meet, else the inline control the selection is in. Inside an inline control a paste is read as one sentence — the fragment's paragraphs and block controls, or the plain text's lines, join with a space — and a fragment holding a table is refused there with `fragment-not-inline`. A copy carries its last paragraph mark only with the whole paragraph, as Word's smart paragraph selection does, so the end of a paragraph pasted mid-paragraph no longer divides the host. A host can set `setPastePolicy` to rewrite the blocks a rich paste lands and to refuse what it would publish, with no lower-fidelity flavour retried; and `ContentControlPromptDisplay.press: 'caret'` meets a prompt with the caret at its start instead of selecting it whole.
