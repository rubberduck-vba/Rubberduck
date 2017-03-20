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





