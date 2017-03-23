Attribute VB_Name = "FakeTests"
Option Explicit

Option OnPrivate Module

'HERE BE DRAGONS.  Save your work in ALL open windows.
'@TestModule
'@Folder("Tests")

Private Assert As New Rubberduck.AssertClass
Private Fakes As New Rubberduck.FakesProvider

'@TestMethod
Public Sub InputBoxFakeWorks()
    On Error GoTo TestFail
    
    Dim userInput As String
    With Fakes.InputBox
        .Returns vbNullString, 1
        .ReturnsWhen "Prompt", "Second", "User entry 2", 2
        userInput = InputBox("First")
        Assert.IsTrue userInput = vbNullString
        userInput = InputBox("Second")
        Assert.IsTrue userInput = "User entry 2"
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub MsgBoxFakeWorks()
    On Error GoTo TestFail
    
    With Fakes.MsgBox
        .Returns vbOK
        MsgBox "This is faked"
        .Verify.Once
        .Verify.Parameter "Prompt", "This is faked"
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TimerFakeWorks()
    On Error GoTo TestFail

    With Fakes.Timer
        .Returns 1234
        Debug.Print Timer
        .Verify.Once
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TimerFakePassThroughWorks()
    On Error GoTo TestFail

    With Fakes.Timer
        .PassThrough = True
        Debug.Print Timer
        .Verify.Once
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub DoEventsFakeWorks()
    On Error GoTo TestFail

    With Fakes.DoEvents
        DoEvents
        .Verify.Once
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub DoEventsFakePassThroughWorks()
    On Error GoTo TestFail

    With Fakes.DoEvents
        .PassThrough = True
        DoEvents
        .Verify.Once
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub ShellFakeWorks()
    On Error GoTo TestFail

    With Fakes.Shell
        .Returns 666.666
        Shell "C:\Windows\notepad.exe"
        .Verify.Once
        .Verify.Parameter "PathName", "C:\Windows\notepad.exe"
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub ShellFakePassThroughWorks()
    On Error GoTo TestFail

    With Fakes.Shell
        .PassThrough = True
        Shell "C:\Windows\notepad.exe"
        .Verify.Once
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

''@TestMethod
'Public Sub CurDirFakeNoArgsWorks()
'    On Error GoTo TestFail

'    With Fakes.CurDir
'        .Returns "C:\Foo"
'        Debug.Print CurDir
'        .Verify.Once
'    End With

'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'End Sub

''@TestMethod
'Public Sub CurDirFakeWorks()
'    On Error GoTo TestFail

'    With Fakes.CurDir
'        .Returns "C:\Foo"
'        Debug.Print CurDir("C")
'        .Verify.Once
'        .Verify.Parameter "Drive", "C"
'    End With

'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'End Sub
