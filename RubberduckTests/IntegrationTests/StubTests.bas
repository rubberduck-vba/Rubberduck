Attribute VB_Name = "StubTests"
Option Explicit

Option Private Module

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

''@TestMethod
'Public Sub RmDirStubPassThroughWorks()
'    On Error GoTo TestFail
'
'    With Fakes.RmDir
'        .PassThrough = True
'        RmDir "C:\Test\Foo"
'        .Verify.Once
'    End With
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.Description
'End Sub

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
        ChDrive "C"
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

'@TestMethod
Public Sub SaveSettingStubWorks()
    On Error GoTo TestFail

    With Fakes.SaveSetting
        SaveSetting "MyApp", "MySection", "MyKey", "MySetting"
        .Verify.Once
        .Verify.Parameter "appname", "MyApp"
        .Verify.Parameter "section", "MySection"
        .Verify.Parameter "key", "MyKey"
        .Verify.Parameter "setting", "MySetting"
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub SaveSettingStubPassThroughWorks()
    On Error GoTo TestFail

    With Fakes.SaveSetting
        .PassThrough = True
        SaveSetting "MyApp", "MySection", "MyKey", "MySetting"
        .Verify.Once
        Assert.IsTrue GetSetting("MyApp", "MySection", "MyKey") = "MySetting"
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub DeleteSettingStubWorks()
    On Error GoTo TestFail

    With Fakes.DeleteSetting
        DeleteSetting "MyApp", "MySection", "MyKey"
        .Verify.Once
        .Verify.Parameter "appname", "MyApp"
        .Verify.Parameter "section", "MySection"
        .Verify.Parameter "key", "MyKey"
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub DeleteSettingStubPassThroughWorks()
    On Error GoTo TestFail

    With Fakes.DeleteSetting
        .PassThrough = True
        DeleteSetting "MyApp", "MySection", "MyKey"
        .Verify.Once
        Assert.IsTrue GetSetting("MyApp", "MySection", "MyKey") = vbNullString
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub RandomizeStubWorks()
    On Error GoTo TestFail

    With Fakes.Randomize
        Randomize 0.5
        .Verify.Once
        .Verify.Parameter "number", 0.5
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub RandomizeStubPassThroughWorks()
    On Error GoTo TestFail

    With Fakes.Randomize
        .PassThrough = True
        Randomize
        .Verify.Once
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub SetAttrFakeWorks()
    On Error GoTo TestFail

    With Fakes.SetAttr
        SetAttr "C:\Test\dummy.txt", vbHidden + vbReadOnly
        .Verify.Once
        .Verify.Parameter "pathname", "C:\Test\dummy.txt"
        .Verify.Parameter "attributes", CInt(vbHidden + vbReadOnly)
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub SetAttrFakePassThroughWorks()
    On Error GoTo TestFail

    With Fakes.SetAttr
        .PassThrough = True
        SetAttr "C:\Test\dummy.txt", vbHidden + vbReadOnly
        .Verify.Once
        .Verify.Parameter "pathname", "C:\Test\dummy.txt"
        .Verify.Parameter "attributes", CInt(vbHidden + vbReadOnly)
        Assert.IsTrue GetAttr("C:\Test\dummy.txt") = vbHidden + vbReadOnly
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub FileCopyFakeWorks()
    On Error GoTo TestFail

    With Fakes.FileCopy
        FileCopy "C:\Test\dummy.txt", "C:\Test\copied.txt"
        .Verify.Once
        .Verify.Parameter "oldpathname", "C:\Test\dummy.txt"
        .Verify.Parameter "newpathname", "C:\Test\copied.txt"
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub FileCopyFakePassThroughWorks()
    On Error GoTo TestFail

    With Fakes.FileCopy
        .PassThrough = True
        FileCopy "C:\Test\dummy.txt", "C:\Test\copied.txt"
        .Verify.Once
        .Verify.Parameter "oldpathname", "C:\Test\dummy.txt"
        .Verify.Parameter "newpathname", "C:\Test\copied.txt"
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

