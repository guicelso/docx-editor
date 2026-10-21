---
'@docx-editor.dev/core': minor
---

An atom is one unit with its tags: the caret meets it before or after it and never inside, also where the host draws a tag on it. A slot beside an atom's tag stands outside the atom — typing there lands before or after it — and where no tag is drawn an atom's edge decides the same. A key at an atom's outer edge selects it whole with or without tags, as a press does, an empty atom included, and the next key removes it. An atom's tags name their control (`data-piece-control`) and carry its `data-hover` and `data-selected`.

A piece's names gain `role`, published as `data-piece-role`: the part a piece plays in the structure around it, apart from its state. The neutral chip is painted in the weight layout measured its label in.
