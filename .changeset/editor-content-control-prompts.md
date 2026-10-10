---
'@docx-editor.dev/pro': patch
---

`setContentControlPrompts` shows a host's text in each content control that holds its placeholder, in the placeholder's style and as one unit. The document, the saved bytes and the clipboard keep the stored placeholder, and a changed prompt lays out again only the paragraph that shows it. `ContentControlTagSubject` is now `ContentControlSubject`, shared by tags and prompts.
