#Every programmer needs a Rubberduck.

![Rubberduck](http://i.stack.imgur.com/taIMg.png)

Rubberduck is a COM Add-In for the VBA IDE that makes VBA development even more enjoyable, by extending the Visual Basic Editor (VBE) with menus, toolbars and toolwindows that enable things we didn't even think were possible when we first started this project.

If you're learning VBA, Rubberduck can help you avoid a few common beginner mistakes, and can probably show you a trick or two - even if you're only ever writing *macros*. If you're a more advanced programmer, you will appreciate the richness of Rubberduck's feature set.

[**Follow us on Twitter!**](https://twitter.com/rubberduckvba)


---

#Features

##Code Explorer

The VBE's *Project Explorer* was nice... in 1999. Get the same bird's eye view of your project and navigate anywhere, with the *Code Explorer* dockable toolwindow:

![Code Explorer Window](http://i.stack.imgur.com/yilHM.png)

This tree view drills down to *member* level, so not only you can see modules with their properties, functions and procedures, you also get to see a module's fields, constants, enums (and their members) and user-defined types (and their members) - without having to bring up the *object browser*.

##To-do Items

Ever wish you had a task list built into the VBA IDE? You don't have to wish anymore: it's here! Rubberduck searches your code for `TODO:` comments (or whatever you configure as "todo" markers) and displays them all in a convenient dockable toolwindow. Double-click on an item in the list to navigate to that location in the code.

![Todo Explorer dockable toolwindow](http://imgur.com/Xl1hfcQ.png)

Rubberduck comes with default markers and priority levels, but that's 100% configurable.

##Test Explorer

Fully integrated unit testing, with zero boiler plate code (a little comment doesn't really count, right?). Use late-binding to create a `Rubberduk.AssertClass` object, or let Rubberduck automatically add a reference to its type library, and start writing unit tests for your VBA code:

    '@TestModule
    Private Assert As New Rubberduck.AssertClass
    
    '@TestMethod
    Public Sub MyMethodReturnsTrue()
        Assert.IsTrue MyMethod
    End Sub
    
    Public Sub TestReferenceEquals()
        Dim collection1 As New Collection
        Dim collection2 As Collection
        Set collection2 = collection1
        Assert.AreSame collection1, collection2
    End Sub

The `'@TestModule` marker is merely a hint to tell Rubberduck to search that module for test methods; the `'@TestMethod` marker isn't needed if the method starts with the word `Test`. 
The *Test Explorer* dockable toolwindow lists all tests found in all opened VBProjects, and makes it easy to add new test modules and methods - and run them:

![Test Explorer dockable toolwindow](http://i.stack.imgur.com/gOMfO.png)

##Code Inspections

Find various code issues in your code - and fix them, with just a few clicks! 

![Code Inspections dockable toolwindow](http://i.imgur.com/djvt8H5.png)

In the event where you would have too many docked windows, Rubberduck offers you a toolbar to quickly navigate and fix code issues:

![Code Inspections toolbar](http://i.stack.imgur.com/0MSot.png)

---

#Coming Up

The following features are currently under development and scheduled for the next version (v1.2):

###ANTLR-Powered Parser: More Code Inspections

Rubberduck's most powerful features will make extensive use of parse trees. We are currently working on re-implementing everything that needs a parser, to open the door to deeper code analysis, and... refactorings:

![extract method](http://i.stack.imgur.com/FhUwt.png)

More refactorings are planned for v1.3; see our [milestones](https://github.com/retailcoder/Rubberduck/milestones) for all the details.

###GitHub/Source Control Integration

This feature will make it possible to push your VBA code to your GitHub repository in separate code files, and to pull commits into the IDE, straight from the IDE. Read that again if you don't believe it, it's real.

---

#Contributing

If you're a C# developer looking for a fun project to contribute to, feel free to fork the project and 
[come meet the devs in Code Review's "VBA" chatroom](http://chat.stackexchange.com/rooms/14929/vba) - we'll be happy to answer your questions and help you help us take the VBE into the 21st century!

Some issues are tagged with [help-wanted](https://github.com/retailcoder/Rubberduck/labels/help-wanted), but that doesn't mean we can't use some help with anything else in the project - if this project interests you, we want to hear from you!

---

#Installation

Visit our releases page, [download the installer](https://github.com/retailcoder/Rubberduck/releases/tag/v1.1), and run the Setup.exe.

Please note that this software has only been tested on Office 2007 & 2010.
Please feel free to test it on other versions and [submit any bugs on our issue tracker](https://github.com/retailcoder/Rubberduck/issues).

If you're **upgrading** from a previous version, you will need to completely uninstall it before installing the newest release. Be sure to back up the `rubberduck.config` file in the `\AppData\Roaming\Rubberduck\` directory prior to installation.

---

#License

The MIT License (MIT)

Copyright (c) 2014 Mathieu Guindon & Christopher McClellan

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

---

#Attributions

##[ANTLR](http://www.antlr.org/)

As of v1.2, Rubberduck is empowered by the awesomeness of ANTLR.

> **What is ANTLR?**

> *ANTLR (ANother Tool for Language Recognition) is a powerful parser generator for reading, processing, executing, or translating structured text or binary files. It's widely used to build languages, tools, and frameworks. From a grammar, ANTLR generates a parser that can build and walk parse trees.*

We're not doing half of what we could be doing with this amazing tool. Try it, see for yourself!

#Icons

We didn't come up with these icons ourselves! Here's who did what:

##[Fugue Icons](http://p.yusukekamiyamane.com/)

This beautiful suite of professional-grade icons packs over 3,570 icons (16x16). You name it, there's an icon for that.

> (C) 2012 Yusuke Kamiyamane. All rights reserved. 
These icons are licensed under a [Creative Commons Attribution 3.0 License](http://creativecommons.org/licenses/by/3.0/).
If you can't or don't want to provide attribution, please [purchase a royalty-free license](http://p.yusukekamiyamane.com/).

##[Microsoft Visual Studio Image Library](http://www.microsoft.com/en-ca/download/details.aspx?id=35825)

Icons in the `./Resources/Microsoft/` directory are licensed under Microsoft's Software License Terms, must be used accordingly with their meaning / file name.

> You have a right to Use and Distribute these files. This means that you are free to copy and use these images in documents and projects that you create, but you may not modify them in anyway.

For more information, please see the EULAs in the [./Resources/Microsoft/ directory](https://github.com/retailcoder/Rubberduck/tree/master/RetailCoder.VBE/Resources/Microsoft).

 * [Visual Studio 2013 Image Library EULA](https://github.com/retailcoder/Rubberduck/blob/master/RetailCoder.VBE/Resources/Microsoft/Visual%20Studio%202013%20Image%20Library%20EULA.rtf)
 * [Visual Studio 2012 Image Library EULA](https://github.com/retailcoder/Rubberduck/blob/master/RetailCoder.VBE/Resources/Microsoft/Visual%20Studio%202012%20Image%20Library%20EULA.rtf)
