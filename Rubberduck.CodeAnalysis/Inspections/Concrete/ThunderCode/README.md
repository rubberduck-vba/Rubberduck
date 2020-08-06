## Rubberduck.CodeAnalysis.Inspections.Concrete.ThunderCode

Inspections in this namespace should have a *hidden* attribute on the summary tag, e.g. `<summary hidden="true">`. When the xml-doc is processed, this attribute toggles this content to be displayed:

> *This feature is hidden. It could be an Easter egg, or a problematic feature that is likely (hopefully?) disabled by default.*
