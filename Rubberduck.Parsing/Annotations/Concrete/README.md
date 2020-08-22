## Rubberduck.Parsing.Annotations.Concrete

All concrete annotation implementations (classes) must be in their own .cs source file in this namespace/folder; avoid including xml-doc on classes in this namespace that aren't annotations.

The xml-doc content in this namespace is automatically downloaded, processed, and ultimately served on the rubberduckvba.com website feature pages.

Examples for attribute annotations should include a `<before>` tag for each module showing the **code pane** code, and an `<after>` tag containing the exported code showing the hidden attribute(s).
Each annotation can have as many examples using as many modules of as many types as necessary. The following string values are recognized as module types:

- "Standard Module"
- "Class Module"
- "Predeclared Class"
- "Interface Module"
- "UserForm Module"
- "Document Module"

The "edit this page" link on each page generated from xml-doc content in this namespace, links to `https://github.com/rubberduck-vba/Rubberduck/edit/next/Rubberduck.Parsing/Annotations/Concrete/{annotation-name}.cs`; it is imperative that the files' folder location corresponds to their namespace, lest we generate broken links on the website.

The content generated from xml-doc in this namespace (and any concrete inspections in a namespace under it) is accessible at `https://rubberduckvba.com/annotations/details/{annotation-name}`.