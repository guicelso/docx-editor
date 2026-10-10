---
'@docx-editor.dev/pro': patch
---

`mergeFragmentIntoPackage` takes a style policy. `styles: 'source'` (the default, and what a paste uses) keeps the fragment's look: a style whose definition differs from the target's same-id style is imported under a fresh id, and the fragment's defaults are stamped where the target's would re-resolve. `styles: 'destination'` is Word's "Use Destination Styles", the way a building block lands in a document: a style the target has by name and type is the target's — whatever its localized id —, a style it lacks is imported, and nothing is stamped, so the content takes the target's look.
