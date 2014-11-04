RetailCoderVBE
==============

A COM Add-In for the VBA IDE.

##Registry Keys

The GUID and ProgId for the `RetailCoderVBE.Extension` class must be registered in the Windows Registry. There is no automated deployment process in place at this point, so the keys must be added & configured manually with RegEdit.

Should there already be a CLSID with the same GUID, a new value will need to be generated and the code recompiled with the new GUID, before the add-in can run.

###64 bit Office

    [HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins64\RetailCoderVBE]
     ~> [CommandLineSafe] (DWORD:00000000)
     ~> [Description] ("RetailCoderVBE add-in for VBA IDE.")
     ~> [LoadBehavior] (DWORD:00000003)
     ~> [FriendlyName] ("RetailCoderVBE")
   
    [HKEY_CLASSES_ROOT\CLSID\{8D052AD8-BBD2-4C59-8DEC-F697CA1F8A66}]
     ~> [@] ("RetailCoderVBE.Extension")
     
    [HKEY_CLASSES_ROOT\CLSID\{8D052AD8-BBD2-4C59-8DEC-F697CA1F8A66}\InprocServer32]
     ~> [@] ("mscoree.dll")
     ~> [ThreadingModel] ("Both")
     ~> [Class] ("RetailCoderVBE.Extension")
     ~> [Assembly] ("RetailCoderVBE")
     ~> [RuntimeVersion] ("v2.0.50727")
     ~> [CodeBase] ("file:///C:\Dev\RetailCoder\RetailCoder.VBE\RetailCoder.VBE\bin\Debug\RetailCoderVBE.dll")
   
    [HKEY_CLASSES_ROOT\CLSID\{8D052AD8-BBD2-4C59-8DEC-F697CA1F8A66}\InprocServer32]
     ~> [@] ("RetailCoderVBE.Extension")

###32 bit Office

In the root folder of the project, there is a *.reg file that will install the necessary registry keys.

	RetailCoderVBE32.reg

	 
##Features

The add-in inserts a [Test] menu to the VBA IDE main menu bar, as well as a [Test] commandbar with the following buttons:
   - **Run All Tests** finds all test methods in all opened projects, and runs them, then displays the Test Explorer.
   - **Test Explorer** displays a stays-on-top sizeable window featuring a grid that lists all test methods in all opened projects, along with their last result.

---

**Test Explorer**

The *Test Explorer* allows browsing/finding, running, and adding unit tests to the active VBProject:

![Test Explorer window](http://i.imgur.com/iNpuaJR.png)

The **Refresh** command synchronizes the test methods with the code in the IDE, but if test methods are added from within the Text Explorer then the new tests will appear automatically.

The **Run** menu makes running the tests as convenient as in the .NET versions of Visual Studio:

![Test Explorer 'Run' menu](http://i.imgur.com/cC6cYGg.png)

"Selected Tests" refer to the selection in the grid, not in the IDE.

The **Add** menu makes it easy to add new tests:

![Test Explorer 'Add' menu](http://i.imgur.com/6mFRlQE.png)

Adding a *Test Module* ensures the active VBProject has a reference to the add-in's type library, then adds a new standard code module with this content:

    '@TestModule
    Option Explicit
    Private Assert As New RetailCoderVBE.AssertClass

Adding a *Test Method* adds this template snippet at the end of the active test module:

    '@TestMethod
    Public Sub TestMethod1() 'TODO: Rename test
        On Error GoTo TestFail
    
        'Arrange
    
        'Act

        'Assert
        Assert.Inconclusive
        
    TestExit:
        Exit Sub
    TestFail:
        If Err.Number <> 0 Then
            Assert.Fail "Test raised an error: " & Err.Description
        End If
        Resume TestExit
    End Sub
    
Adding a *Test Method (expected error)* adds this template snippet at the end of the active test module:

    '@TestMethod
    Public Sub TestMethod2() 'TODO: Rename test
        Const ExpectedError As Long = 0 'TODO: Change to expected error number
        On Error GoTo TestFail
        
        'Arrange
    
        'Act
    
        'Assert
        Assert.Fail "Expected error was not raised."
        
    TestExit:
        Exit Sub
    TestFail:
        Assert.AreEqual ExpectedError, Err.Number
        Resume TestExit
    End Sub

The number at the end of the generated method name depends on the number of test methods in the test module.
