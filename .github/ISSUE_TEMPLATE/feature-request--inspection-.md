---
name: Feature request (inspection)
about: Suggest something Rubberduck could find in user code and warn about
title: ''
labels: enhancement, feature-inspections, up-for-grabs
assignees: ''

---

**What**
Describe what the new inspection should find in the user's VBA code; identify the type of inspection: is it about code quality (e.g. potential bugs), is it more of a language opportunity (e.g. obsolete statements), a performance opportunity (e.g. iterating an array with a `For Each` loop), or a Rubberduck opportunity (e.g. something Rubberduck can do or help with, but needs the user code to be modified a bit).

**Why**
Describe the rationale behind this inspection - *why* finding what we're looking for in the user's code is noteworthy: justify the inspection.

**Example**
This code should trigger the inspection:

```vb
Public Sub DoSomething()
    '...
End Sub
```

---

**QuickFixes**
Should Rubberduck offer one or more quickfix(es) for this inspection? Describe them here (note: all inspections allow for `IgnoreOnceQuickFix`, unless explicitly specified):

1. **QuickFix Name**

    Example code, after quickfix is applied:

    ```vb
    Public Sub DoSomething()
        '...
    End Sub
    ```

---

**Resources**
Each inspection needs a number of resource strings - please provide a suggestion here:

 - **InspectionNames**: the name of the inspection, as it appears in the inspection settings.
 - **InspectionInfo**: the detailed rationale for the inspection, as it appears in the inspection results toolwindow's bottom panel.
 - **InspectionResults**: the resource string for an inspection result; if the string needs placeholders, identify what goes in for them (e.g. `Variable {0} is not used`, {0}:the name of the variable).
