![Rubberduck](http://i.stack.imgur.com/taIMg.png)

Rubberduck is a COM Add-In for the VBA IDE that makes VBA development even more enjoyable. 

##Features:
###Unit testing

Fully integrated unit testing with minimal (read: next to none) boiler plate code. Just add a reference to Rubberduck and create a new module-scoped `AssertClass` instance and you're ready to start writing tests. 

    '@TestModule
    Private Assert As New Rubberduck.AssertClass
    
    
    '@TestMethod
    Public Sub OnePlusOneIsTwo()
        Const expected As Long = 2
        
        Assert.AreEqual expected, Add(1, 1)
    End Sub

Rubberduck will find the Module and Procedure attributes and display your test methods for you in the Test Explorer.

![Test Explorer Window](http://imgur.com/NepssQ8.png)

###To-do items

Ever wish you had a task list built into the VBA IDE? You don't have to wish anymore. It's here. Rubberduck searches your code for `TODO:` comments and displays them all in one convenient location. Double-click on an item in the Todo List and jump to that location in the code. 

![Todo List Window](http://imgur.com/Xl1hfcQ.png)

The comments are also configurable, so you can decide what comments to add to your Todo List and what Priority level they should be.

###Code Explorer

Get a bird's eye view of your project and navigate anywhere, with the Code Explorer.

![Code Explorer Window](http://i.imgur.com/bkDOB4w.png)

###Code Inspections

Find code issues in your code - and fix them, with just a few clicks!

![Code Inspections window](http://i.imgur.com/djvt8H5.png)

##Coming Up

The following features are currently under development:

###ANTLR-Powered Parser

Rubberduck's most powerful features will make extensive use of parse trees. We are currently working on re-implementing everything that needs a parser, to open the door to deeper code analysis, and... refactorings:

![extract method](http://i.stack.imgur.com/FhUwt.png)

###GitHub/Source Control Integration

This feature will make it possible to push your VBA code to your GitHub repository in separate code files, and to pull commits into the IDE, straight from the IDE.

###Rubberduck.Reflection

At first this COM-visible type library will *simply* let you write some VBA meta-code, that can iterate all opened projects and every code module a bit like the VBE API does, except you will also be able to iterate enum members, user-defined types, fields, properties, methods, functions, and every declared constant, variable, external function... without even needing to enable access to the VBE API in the macro security settings.

---

##Installation 
Visit our releases page, [download the installer](https://github.com/retailcoder/Rubberduck/releases/tag/v1.01-alpha2), and run the Setup.exe. 

Please note that this software has only been tested on Office 2007 & 2010.
Please feel free to test it on other versions and [submit any bugs on our issue tracker](https://github.com/retailcoder/Rubberduck/issues).

If you're **upgrading** from a previous version, you will need to completely uninstall it before installing the newest release. Be sure to back up the `rubberduck.config` file in the `\AppData\Roaming\Rubberduck\` directory prior to installation.

---

##Contributing

[Come meet the devs in Code Review's "VBA" chatroom](http://chat.stackexchange.com/rooms/14929/vba)!

---

## License

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

##Icons attribution

###Microsoft Visual Studio Image Library

The image files in the ./Resources/Microsoft/ directory are licensed under Microsoft's Software License Terms.

You have a right to Use and Distribute these files. This means that you are free to copy and use these images in documents and projects that you create, but you may not modify them in anyway. For more information, please see the EULAs in the [./Resources/Microsoft/ directory](https://github.com/retailcoder/Rubberduck/tree/master/RetailCoder.VBE/Resources/Microsoft).

 * [Visual Studio 2013 Image Library EULA](https://github.com/retailcoder/Rubberduck/blob/master/RetailCoder.VBE/Resources/Microsoft/Visual%20Studio%202013%20Image%20Library%20EULA.rtf)
 * [Visual Studio 2012 Image Library EULA](https://github.com/retailcoder/Rubberduck/blob/master/RetailCoder.VBE/Resources/Microsoft/Visual%20Studio%202012%20Image%20Library%20EULA.rtf)
