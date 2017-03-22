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