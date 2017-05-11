<img src="http://i.stack.imgur.com/vmqXM.png" width="320" />

Branch     | Description | Build Status |
|------------|---|--------------|
| **master** | The last released build | ![master branch build status][masterBuildStatus] |
| **next**   | The current build (dev)  | ![next branch build status][nextBuildStatus] |

[nextBuildStatus]:https://ci.appveyor.com/api/projects/status/we3pdnkeebo4nlck/branch/next?svg=true
[masterBuildStatus]:https://ci.appveyor.com/api/projects/status/we3pdnkeebo4nlck/branch/master?svg=true

[![Average time to resolve an issue](http://isitmaintained.com/badge/resolution/rubberduck-vba/rubberduck.svg)](http://isitmaintained.com/project/rubberduck-vba/rubberduck "Average time to resolve an issue") [![Percentage of issues still open](http://isitmaintained.com/badge/open/rubberduck-vba/rubberduck.svg)](http://isitmaintained.com/project/rubberduck-vba/rubberduck "Percentage of issues still open") 

> **[rubberduckvba.com](http://rubberduckvba.com)** [Wiki](https://github.com/retailcoder/Rubberduck/wiki) [Rubberduck News](https://rubberduckvba.wordpress.com/) 
> contact@rubberduckvba.com  
> Follow [@rubberduckvba](https://twitter.com/rubberduckvba) on Twitter 

---

## What is Rubberduck?

It's an add-in for the VBA IDE, the glorious *Visual Basic Editor* (VBE) - which hasn't seen an update in this century, but that's still in use everywhere around the world. Rubberduck wants to give its users access to features you would find in the VBE if it had kept up with the features of Visual Studio and other IDE's in the past, oh, *decade* or so.

Rubberduck wants to help its users write better, cleaner, maintainable code. The many **code inspections** and **refactoring tools** help harmlessly making changes to the code, and **unit testing** helps writing a *safety net* that makes it easy to know exactly what broke when you made that *small little harmless modification*.

Rubberduck wants to bring VBA into the 21st century, and wants to see more open-source VBA repositories on [GitHub](https://github.com/) - VBA code and **source control** don't traditionally exactly work hand in hand; unless you've automated it, exporting each module one by one to your local repository, fetching the remote changes, re-importing every module one by one back into the project, ...is *a little bit* tedious. Rubberduck integrates Git into the IDE, and handles all the file handling behind the scenes - a bit like Visual Studio's *Team Explorer*.

---

If you're learning VBA, Rubberduck can help you avoid a few common beginner mistakes, and can probably show you a trick or two - even if you're only ever writing *macros*. If you're a more advanced programmer, you will appreciate the richness of [Rubberduck's feature set](https://github.com/retailcoder/Rubberduck/wiki/Features). See the [Installing](https://github.com/rubberduck-vba/Rubberduck/wiki/Installing) wiki page.

If you're a C# developer looking for a fun project to contribute to, see the [Contributing](https://github.com/rubberduck-vba/Rubberduck/wiki/Contributing) wiki page.

---

## License

Rubberduck is a COM add-in for the VBA IDE (VBE).

Copyright (C) 2014-2016 Mathieu Guindon & Christopher McClellan

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see http://www.gnu.org/licenses/.

---

# Attributions

## Software & Libraries

### [Ninject](http://www.ninject.org)

**What is Ninject?**

> *Ninject is a lightning-fast, ultra-lightweight dependency injector for .NET applications. It helps you split your application into a collection of loosely-coupled, highly-cohesive pieces, and then glue them back together in a flexible manner. By using Ninject to support your software's architecture, your code will become easier to write, reuse, test, and modify.*

If you're into *Dependency Injection* and *Inversion of Control*, you'll appreciate Ninject's simple configuration API. Other than in the very specialized startup code, Ninject seems absent of Rubberduck's code base, which indeed knows nothing of Ninject or any IoC framework: don't let it fool you - Ninject is responsible for instantiating essentially.. every single class in the solution.

Rubberduck uses the following Ninject [extensions](http://www.ninject.org/extensions):

 - Ninject.Extensions.Conventions
 - Ninject.Extensions.Factory
 - Ninject.Extensions.Interception
 - Ninject.Extensions.Interception.DynamicProxy
 - Ninject.Extensions.NamedScope

### [ANTLR](http://www.antlr.org/)

As of v1.2, Rubberduck is empowered by the awesomeness of ANTLR.

> **What is ANTLR?**

> *ANTLR (ANother Tool for Language Recognition) is a powerful parser generator for reading, processing, executing, or translating structured text or binary files. It's widely used to build languages, tools, and frameworks. From a grammar, ANTLR generates a parser that can build and walk parse trees.*

We're not doing half of what we could be doing with this amazing tool. Try it, see for yourself!

### [LibGit2Sharp](https://github.com/libgit2/libgit2sharp)

**What is LibGit2Sharp?**

LibGit2Sharp is the library that has allowed us to integrate Git right into the VBA IDE (and as a nice bonus, expose a nice API that handles the nitty gritty of importing source files to and from the IDE to a repo for you).

> LibGit2Sharp brings all the might and speed of libgit2, a native Git implementation, to the managed world of .Net and Mono.

**Okay, so what is [libgit2](https://libgit2.github.com/)?**

> libgit2 is a portable, pure C implementation of the Git core methods provided as a re-entrant linkable library with a solid API, allowing you to write native speed custom Git applications in any language which supports C bindings.

Which basically means it's a reimplementation of Git in C. It also [happens to be the technology Microsoft uses for their own Git integration with Visual Studio](http://www.hanselman.com/blog/GitSupportForVisualStudioGitTFSAndVSPutIntoContext.aspx).

### [AvalonEdit](http://avalonedit.net)

Source code looks a lot better with syntax highlighting, and AvalonEdit excels at it. 

> AvalonEdit is a WPF-based text editor component. It was written by [Daniel Grunwald](https://github.com/dgrunwald) for the [SharpDevelop](http://www.icsharpcode.net/OpenSource/SD/) IDE. Starting with version 5.0, AvalonEdit is released under the [MIT license](http://opensource.org/licenses/MIT).

We're currently only using a tiny bit of this code editor's functionality (more to come!).

### [EasyHook](http://easyhook.github.io/index.html)

Without the EasyHook library, many of our more advanced Unit Testing features would simply not be possible.  This library really lives up to its name, and allows us to intercept and inspect traffic through VBE7.dll and other unmanged libraries.
 
> EasyHook makes it possible to extend (via hooking) unmanaged code APIs with pure managed functions, from within a fully managed environment on 32- or 64-bit Windows XP SP2, Windows Vista x64, Windows Server 2008 x64, Windows 7, Windows 8.1, and Windows 10. 

EasyHook is released under the [MIT license](https://github.com/EasyHook/EasyHook#license).

### [WPF Localization Using RESX Files](http://www.codeproject.com/Articles/35159/WPF-Localization-Using-RESX-Files)

This library makes localizing WPF applications at runtime using resx files a breeze. Thank you [Grant Frisken](http://www.codeproject.com/script/Membership/View.aspx?mid=1079060)!

> Licensed under [The Code Project Open License](http://www.codeproject.com/info/cpol10.aspx) with the [author's permission](http://www.codeproject.com/Messages/5272045/Re-License.aspx) to re-release under the GPLv3.

## Icons

We didn't come up with these icons ourselves! Here's who did what:

### [Fugue Icons](http://p.yusukekamiyamane.com/)

This beautiful suite of professional-grade icons packs over 3,570 icons (16x16). You name it, there's an icon for that.

> (C) 2012 Yusuke Kamiyamane. All rights reserved. 
These icons are licensed under a [Creative Commons Attribution 3.0 License](http://creativecommons.org/licenses/by/3.0/).
If you can't or don't want to provide attribution, please [purchase a royalty-free license](http://p.yusukekamiyamane.com/).

### [SharpDevelop](https://github.com/icsharpcode/SharpDevelop.git)

Icons in the `./Resources/Custom/` directory were created by (or modified using elements from) the SharpDevelop icon set licensed under the [MIT license](https://opensource.org/licenses/MIT).

---

## [JetBrains](https://www.jetbrains.com) | [ReSharper](https://www.jetbrains.com/resharper/)

[![JetBrains ReSharper logo](https://cloud.githubusercontent.com/assets/5751684/20271309/616bb740-aa58-11e6-91c9-65287b740985.png)](https://www.jetbrains.com/resharper/)

Since the project's early days, JetBrains' Open-Source team has been supporting Rubberduck - and we deeply thank them for that. ReSharper has been not only a tool we couldn't do without; it's been an inspiration, the ultimate level of polished perfection to strive for in our own IDE add-in project. So just like you're missing out if you write VBA and you're not using Rubberduck, you're missing out if you write C# and aren't using ReSharper.

<sub>Note: Rubberduck is not a JetBrains product. JetBrains does not contribute and is not affiliated to the Rubberduck project in any way.</sub>

---

# Overview

The first thing you will notice of Rubberduck is its commandbar and menus; Rubberduck becomes part of the VBE, but at startup you'll notice almost everything is disabled, and the Rubberduck commandbar says "Pending":

![A 'Refresh' button, and 'Pending' state label in the Rubberduck commandbar](https://cloud.githubusercontent.com/assets/5751684/21707782/2e5a1a42-d3a0-11e6-87a3-c36ff65f9a79.png)

This button is how Rubberduck keeps in sync with what's in the IDE: when it's Rubberduck itself making changes to the code, it will refresh automatically, but if you make changes to the code and then want to use Rubberduck features, you'll need Rubberduck to *parse* the code first.

The status label will display various steps:

 - **Loading declarations**: Rubberduck noticed new project references and is retrieving information from the COM type libraries.
 - **Parsing**: Rubberduck is creating a parse tree for each new module, and/or updating the parse trees for the modified ones.
 - **Resolving declarations**: The parse trees are being traversed to identify all declarations (variables, procedures, parameters, locals, ...line labels, *everything*).
 - **Resolving references**: The parse trees are being traversed again, this time to locate all identifier references and resolve them all to a specific declaration.
 - **Inspecting**: At this point most features are enabled already; Rubberduck is running its inspections and displaying the results in the *inspection results* toolwindow.
 
That's if everything goes well. Rubberduck assumes the code it's parsing is valid, compilable code that VBA itself can understand.

It's possible you encounter (or write!) code that VBA has no problem with, but that Rubberduck's parser can't handle. When that's the case the Rubberduck commandbar will feature an "error" button:
 
![button tooltip is "1 module(s) failed to parse; click for details."](https://cloud.githubusercontent.com/assets/5751684/21708236/810e9ade-d3a4-11e6-8b4c-c4ec223c066a.png)

Clicking the button brings up a tab in the *Search Results* toolwindow, from which you can double-click to navigate to the exact problematic position in the code:

![Parser errors are all displayed in a "Parser Errors" search results tab](https://cloud.githubusercontent.com/assets/5751684/21708348/86e64b72-d3a5-11e6-9aa8-60cd8d0bec33.png)

The *Code Explorer* will also be displaying the corresponding module node with a red cross icon:

![Module "ThisWorkbook" didn't parse correctly](https://cloud.githubusercontent.com/assets/5751684/21708276/e8f67e50-d3a4-11e6-8c1d-e84d4e9ccce6.png)

You'll find the *Code Explorer* under the *Navigate* menu. By default the Ctrl+R hotkey to display it instead of the VBE's own *Project Explorer*. The treeview lists not only modules, but also every single one of their members, with their signatures if you want. And you can make it arrange your modules into folders, simply by adding a `@Folder("Parent.Child")` annotation/comment to your modules:

![Code Explorer toolwindow](https://cloud.githubusercontent.com/assets/5751684/21708614/90335a46-d3a8-11e6-9e76-61cc3f566c7a.png)

The *inspection results* toolwindow can be displayed by pressing Ctrl+Shift+i (default hotkey), and allows you to double-click to navigate all potential issues that Rubberduck found in your code.

![inspection results](https://cloud.githubusercontent.com/assets/5751684/21708911/d0c47bc4-d3aa-11e6-88f2-b0c9fcfda7ed.png)

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

...with a single click.

---

There's *quite a lot* to Rubberduck, the above is barely even a "quick tour"; the project's [website](http://www.rubberduckvba.com/) lists all the features, and the [wiki](https://github.com/rubberduck-vba/Rubberduck/wiki) will eventually document everything there is to document. Feel free to poke around and break things and [request features / create new issues](https://github.com/rubberduck-vba/Rubberduck/issues/new) too!

0
