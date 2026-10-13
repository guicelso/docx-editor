---
'@docx-editor.dev/react': patch
'@docx-editor.dev/vue': patch
---

A part child for a registry slot the default toolbar does not draw — inserting a table or a footnote, the paragraph marks, a page break — now takes its registry group instead of appending after the arrangement as host content. It sits inside the default group that holds it, in registry order, or in a group of its own after the default ones, and it collapses into "⋯" with that group under the registry label; a bare `<Toolbar.Button slot>` there is the panel's command row. Before, these parts were costed as fixed width and never collapsed, so a host that composed a few entry points without the menu got a row that ran past the bar on a narrow window. A hidden part for a slot the arrangement does not draw adds nothing, and children that drive no slot stay fixed as before.
