![Rubberduck](http://i.stack.imgur.com/taIMg.png)

Rubberduck is a COM Add-In for the VBA IDE that makes VBA development even more enjoyable. 

##Features:
###Unit testing

Fully integrated unit testing with minimal (read: next to none) boiler plate code. Just add a reference to Rubberduck and create a new module scoped `Assert` class and you're ready to start writing tests. 

    '@TestModule
    Private Assert As New Rubberduck.AssertClass
    
    
    '@TestMethod
    Public Sub OnePlusOneIsTwo()
        Const expected As Long = 2
        
        Assert.AreEqual expected, Add(1, 1)
    End Sub

Rubberduck will find the Module and Procedure attributes and display your test methods for you in the Test Explorer.

![Test Explorer Window](http://i.imgur.com/qpCrN30.png)

###To-do items

Ever wish you had a task list built into the VBA IDE? You don't have to wish anymore. It's here. Rubberduck searches your code for `TODO:` comments and displays them all in one convenient location. Double-click on an item in the Todo List and jump to that location in the code. 

![Todo List Window](http://i.stack.imgur.com/3ej9b.png)

The comments are also configurable, so you can decide what comments to add to your Todo List and what Priority level they should be.

###Code Explorer

Get a 30,000 foot view of your project with the Code Explorer.

##Installation 
There is currently not an installer, but we do plan on creating one in the future. In the meantime, [information for installing Rubberduck][install] can be found on our wiki. 
   
[install]:https://github.com/retailcoder/Rubberduck/wiki/Building-Installation
---   

##Icons attribution

###Fugue Icons

(C) 2012 Yusuke Kamiyamane. All rights reserved.
These icons are licensed under a Creative Commons
Attribution 3.0 License.
<http://creativecommons.org/licenses/by/3.0/>

If you can't or don't want to provide attribution, please
purchase a royalty-free license.
<http://p.yusukekamiyamane.com/>

###Microsoft Visual Studio Image Library

The image files in the ./Resources/Microsoft/ directory are licensed under Microsoft's Software License Terms.

You have a right to Use and Distribute these files. This means that you are free to copy and use these images in documents and projects that you create, but you may not modify them in anyway. For more information, please see the EULAs in the [./Resources/Microsoft/ directory](https://github.com/retailcoder/Rubberduck/tree/master/RetailCoder.VBE/Resources/Microsoft).

 * [Visual Studio 2013 Image Library EULA](https://github.com/retailcoder/Rubberduck/blob/master/RetailCoder.VBE/Resources/Microsoft/Visual%20Studio%202013%20Image%20Library%20EULA.rtf)
 * [Visual Studio 2012 Image Library EULA](https://github.com/retailcoder/Rubberduck/blob/master/RetailCoder.VBE/Resources/Microsoft/Visual%20Studio%202012%20Image%20Library%20EULA.rtf)
