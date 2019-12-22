[back to readme.md](https://github.com/rubberduck-vba/Rubberduck/blob/next/README.md)

## Getting Started

<sub>[(installation instructions)](https://github.com/rubberduck-vba/Rubberduck/wiki/Installing)</sub>

### Rubberduck commandbar:
The first thing you will notice of Rubberduck is its commandbar and menus; Rubberduck becomes part of the VBE, but at startup you'll notice almost everything is disabled, and the Rubberduck commandbar says "Pending":

![A 'Refresh' button, and 'Pending' state label in the Rubberduck commandbar](https://cloud.githubusercontent.com/assets/5751684/21707782/2e5a1a42-d3a0-11e6-87a3-c36ff65f9a79.png)

This button is how Rubberduck keeps in sync with what's in the IDE: when it's Rubberduck itself making changes to the code, it will refresh automatically, but if you make changes to the code and then want to use Rubberduck features, you'll need Rubberduck to *parse* the code first.

#### Status Labels:
The status label will display various steps:

 - **Loading declarations**: Rubberduck noticed new project references and is retrieving information from the COM type libraries.
 - **Parsing**: Rubberduck is creating a parse tree for each new module, and/or updating the parse trees for the modified ones.
 - **Resolving declarations**: The parse trees are being traversed to identify all declarations (variables, procedures, parameters, locals, ...line labels, *everything*).
 - **Resolving references**: The parse trees are being traversed again, this time to locate all identifier references and resolve them all to a specific declaration.
 - **Inspecting**: At this point most features are enabled already; Rubberduck is running its inspections and displaying the results in the *inspection results* toolwindow.
 
#### Successful Run:
That's if everything goes well. Rubberduck assumes the code it's parsing is valid, compilable code that VBA itself can understand.

#### Potential Issues:
It's possible you encounter (or write!) code that VBA has no problem with, but that Rubberduck's parser can't handle. When that's the case the Rubberduck commandbar will feature an "error" button:
 
![button tooltip is "1 module(s) failed to parse; click for details."](https://cloud.githubusercontent.com/assets/5751684/21708236/810e9ade-d3a4-11e6-8b4c-c4ec223c066a.png)

#### Issue Resolution Aid:
Clicking the button brings up a tab in the *Search Results* toolwindow, from which you can double-click to navigate to the exact problematic position in the code:

![Parser errors are all displayed in a "Parser Errors" search results tab](https://cloud.githubusercontent.com/assets/5751684/21708348/86e64b72-d3a5-11e6-9aa8-60cd8d0bec33.png).

### Code Explorer:
The *Code Explorer* will also be displaying the corresponding module node with a red cross icon:

![Module "ThisWorkbook" didn't parse correctly](https://cloud.githubusercontent.com/assets/5751684/21708276/e8f67e50-d3a4-11e6-8c1d-e84d4e9ccce6.png)

You'll find the *Code Explorer* under the *Navigate* menu. By default the Ctrl+R hotkey to display it instead of the VBE's own *Project Explorer*. The treeview lists not only modules, but also every single one of their members, with their signatures if you want. And you can make it arrange your modules into folders, simply by adding a `@Folder("Parent.Child")` annotation/comment to your modules:

![Code Explorer toolwindow](https://cloud.githubusercontent.com/assets/5751684/21708614/90335a46-d3a8-11e6-9e76-61cc3f566c7a.png)

### Code Inspections:
The *inspection results* toolwindow can be displayed by pressing Ctrl+Shift+i (default hotkey), and allows you to double-click to navigate all potential issues that Rubberduck found in your code.

![inspection results](https://cloud.githubusercontent.com/assets/5751684/21708911/d0c47bc4-d3aa-11e6-88f2-b0c9fcfda7ed.png)

### Code Smart Indent:
Rubberduck also features a port of the popular "Smart Indenter" add-in (now supports 64-bit hosts, and with a few bugfixes on top of that!), so you can turn this:

````vb
Sub DoSomething()
With ActiveCell
With .Offset(1, 2)
If .value > 100 Then
MsgBox "something"
Else
MsgBox "something else"
End If
End With
End With
End Sub
````

Into this:

````vb
Sub DoSomething()
    With ActiveCell
        With .Offset(1, 2)
            If .value > 100 Then
                MsgBox "something"
            Else
                MsgBox "something else"
            End If
        End With
    End With
End Sub
````

...with a single click!

**Rubberduck > Indent > Current Procedure / Module / Project**  

---
Some helpful shortcuts (these hotkeys are all configurable):

| Shortcut | Function |
| --- | --- |
| `Ctrl + r` | *Rubberduck's*  **Code Explorer** toolwindow, a better version of *VBE's* own **Project Explorer** |
| `Ctrl + Shift + i` | *Rubberduck's* **Code Inspections** toolwindow, to quickly summarize your code issues |
| `Ctrl + p` | *Smart Indent* the current **Procudure** |
| `Ctrl + m` | *Smart Indent* the current **Module** |

---

There's *quite a lot* to Rubberduck, the above is barely even a "quick tour"; the project's [website](http://www.rubberduckvba.com/) lists all the features, and the [wiki](https://github.com/rubberduck-vba/Rubberduck/wiki) will eventually document everything there is to document. Feel free to poke around and break things and [request features / create new issues](https://github.com/rubberduck-vba/Rubberduck/issues/new) too!
