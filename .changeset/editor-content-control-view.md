---
'@docx-editor.dev/pro': patch
---

The editor keeps a host's content-control tags, field tones and field selection across loads, attach and detach, through `setContentControlTags`, `setFieldTones` and `setFieldSelection`. A changed tag lays out again only the paragraph that shows it, so `ContentControlTagDisplay` no longer takes a `token`.
