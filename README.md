RetailCoderVBE
==============

A COM Add-In for the VBA IDE.

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
