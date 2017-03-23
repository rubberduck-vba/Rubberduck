Attribute VB_Name = "StubTests"
Option Explicit On

Option OnPrivate Module

'HERE BE DRAGONS.  Save your work in ALL open windows.
'@TestModule
'@Folder("Tests")

Private Assert As New Rubberduck.AssertClass
Private Fakes As New Rubberduck.FakesProvider

'@TestMethod
Public Sub BeepStubWorks()
    On Error GoTo TestFail
    
    With Fakes.Beep
        Beep
        .Verify.Once
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub BeepStubWorksWithPassthrough()
    On Error GoTo TestFail
    
    With Fakes.Beep
        .PassThrough = True
        Beep
        .Verify.Once
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub SendKeysStubWorks()
    On Error GoTo TestFail

    With Fakes.SendKeys
        SendKeys "{Up}"
        .Verify.Once
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub KillStubWorks()
    On Error GoTo TestFail

    With Fakes.Kill
        Kill "C:\Test\foo.txt"
        .Verify.Once
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

''@TestMethod
'Public Sub KillStubPassThroughWorks()
'    On Error GoTo TestFail
'
'    With Fakes.Kill
'        .PassThrough = True
'        Kill "C:\Test\foo.txt"
'        .Verify.Once
'    End With
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'End Sub

'@TestMethod
Public Sub MkDirStubWorks()
    On Error GoTo TestFail

    With Fakes.MkDir
        MkDir "C:\Test\Foo"
        .Verify.Once
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

''@TestMethod
'Public Sub MkDirStubPassThroughWorks()
'    On Error GoTo TestFail

'    With Fakes.MkDir
'        .PassThrough = True
'        MkDir "C:\Test\Foo"
'        .Verify.Once
'    End With

'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'End Sub

'@TestMethod
Public Sub RmDirStubWorks()
    On Error GoTo TestFail

    With Fakes.RmDir
        RmDir "C:\Test\Foo"
        .Verify.Once
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub RmDirStubPassThroughWorks()
    On Error GoTo TestFail

    With Fakes.RmDir
        .PassThrough = True
        RmDir "C:\Test\Foo"
        .Verify.Once
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub ChDirStubWorks()
    On Error GoTo TestFail

    With Fakes.ChDir
        ChDir "C:\Test"
        .Verify.Once
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub ChDirStubPassThroughWorks()
    On Error GoTo TestFail

    With Fakes.ChDir
        .PassThrough = True
        ChDir "C:\Test"
        .Verify.Once
        Assert.IsTrue CurDir$ = "C:\Test"
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub ChDriveStubWorks()
    On Error GoTo TestFail

    With Fakes.ChDrive
        ChDrive "D"
        .Verify.Once
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub ChDriveStubPassThroughWorks()
    On Error GoTo TestFail

    With Fakes.ChDrive
        .PassThrough = True
        ChDrive "D"
        .Verify.Once
        Assert.IsTrue Left$(CurDir$, 1) = "D"
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub