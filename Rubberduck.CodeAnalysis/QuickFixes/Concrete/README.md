## Rubberduck.CodeAnalysis.QuickFixes.Concrete

All concrete quick-fix implementations (classes) must be in their own .cs source file **in a namespace that corresponds to its folder location**, under this namespace.

The xml-doc content in this namespace is automatically downloaded, processed, and ultimately served on the rubberduckvba.com website feature pages.

Each quick-fix can have as many examples using as many modules of as many types as necessary. The following string values are recognized as module types:

- "Standard Module"
- "Class Module"
- "Predeclared Class"
- "Interface Module"
- "UserForm Module"
- "Document Module"

The "edit this page" link on each page generated from xml-doc content in this namespace, links to `https://github.com/rubberduck-vba/Rubberduck/edit/next/{namespace}/{quickfix-name}.cs`; it is imperative that the files' folder location corresponds to their namespace, lest we generate broken links on the website.

The content generated from xml-doc in this namespace (and any concrete inspections in a namespace under it) is accessible at `https://rubberduckvba.com/quickfixes/details/{quickfix-name}`.