Attribute VB_Name = "FakeTests"
Option Explicit

Option Private Module

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

'@TestMethod
Public Sub EnvironFakeVariantFormWorks()
    On Error GoTo TestFail

    Dim returnVal As Variant
    With Fakes.Environ
        .ReturnsWhen "envstring", "PATH", "C:\Rubberduck", 1
        .ReturnsWhen "envstring", "PATH", "C:\Second", 2
        returnVal = Environ("PATH")
        .Verify.Once
        .Verify.Parameter "envstring", "PATH"
        Assert.IsTrue returnVal = "C:\Rubberduck"
        returnVal = Environ("PATH")
        Assert.IsTrue returnVal = "C:\Second"
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub EnvironFakeStringFormWorks()
    On Error GoTo TestFail

    Dim returnVal As Variant
    With Fakes.Environ
        .ReturnsWhen "envstring", "PATH", "C:\Rubberduck"
        returnVal = Environ$("PATH")
        .Verify.Once
        .Verify.Parameter "envstring", "PATH"
        Assert.IsTrue returnVal = "C:\Rubberduck"
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

Public Sub CurDirFakeNoArgsWorks()
    On Error GoTo TestFail

    Dim returnVal As Variant
    With Fakes.CurDir
        .Returns "C:\Foo"
        returnVal = CurDir()
        .Verify.Once
        Assert.IsTrue returnVal = "C:\Foo"
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub CurDirFakeWorks()
    On Error GoTo TestFail

    Dim returnVal As Variant
    With Fakes.CurDir
        .Returns "C:\Foo"
        returnVal = CurDir("C")
        .Verify.Once
        .Verify.Parameter "Drive", "C"
        Assert.IsTrue returnVal = "C:\Foo"
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub CurDirFakeStringReturnWorks()
    On Error GoTo TestFail

    Dim returnVal As Variant
    With Fakes.CurDir
        .Returns "C:\Foo"
        returnVal = CurDir$("C")
        .Verify.Once
        .Verify.Parameter "Drive", "C"
        Assert.IsTrue returnVal = "C:\Foo"
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub NowFakeWorks()
    On Error GoTo TestFail

    With Fakes.Now
        .Returns #1/1/2018 9:00:00 AM#
        Assert.IsTrue Now = #1/1/2018 9:00:00 AM#
        .Verify.Once
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub NowFakePassThroughWorks()
    On Error GoTo TestFail

    With Fakes.Now
        .Returns #1/1/2018 9:00:00 AM#
        .PassThrough = True
        Assert.IsTrue Now <> #1/1/2018 9:00:00 AM#
        .Verify.Once
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TimeFakeWorks()
    On Error GoTo TestFail

    With Fakes.Time
        .Returns #9:00:00 AM#
        Assert.IsTrue Time = #9:00:00 AM#
        .Verify.Once
    End With
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TimeFakePassThroughWorks()
    On Error GoTo TestFail

    With Fakes.Time
        .Returns #9:00:00 AM#
        .PassThrough = True
        Assert.IsTrue Time <> #9:00:00 AM#
        .Verify.Once
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub DateFakeWorks()
    On Error GoTo TestFail

    With Fakes.Date
        .Returns #1/1/1993#
        Assert.IsTrue Date = #1/1/1993#
        .Verify.Once
    End With
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub DateFakePassThroughWorks()
    On Error GoTo TestFail

    With Fakes.Date
        .Returns #1/1/1993#
        .PassThrough = True
        Assert.IsTrue Date <> #1/1/1993#
        .Verify.Once
    End With
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub MsgBoxAfterInputBoxAnyInvocationFakeWorks()
    On Error GoTo TestFail

    Dim userInput As String

    Fakes.InputBox.ReturnsWhen "Prompt", "Second", "User entry 2"
    Fakes.MsgBox.Returns vbOK

    Dim msgBoxRetVal As Integer
    msgBoxRetVal = MsgBox("This is faked", Title:="My Title")

    Assert.IsTrue msgBoxRetVal = vbOK
    Fakes.MsgBox.Verify.Once

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub InputBoxFakeReturnsWhenWorks()
    On Error GoTo TestFail

    Dim userInput As String
    Fakes.InputBox.ReturnsWhen "prompt", "Dummy1", "dummy1 user input"
    Fakes.InputBox.ReturnsWhen "prompt", "Expected", "expected user input"
    Fakes.InputBox.ReturnsWhen "prompt", "Dummy2", "dummy2 user input"

    userInput = InputBox(prompt:="Expected")

    Assert.AreEqual "expected user input", userInput

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub RndWorks()
    On Error GoTo TestFail

    Dim return1 As Single
    Dim return2 As Single
    With Fakes.Rnd
        .Returns 0.1
        .ReturnsWhen "Number", 1, 0.99

        return1 = Rnd()
        Assert.IsTrue return1 = 0.1

        return2 = Rnd(1)
        Assert.IsTrue return2 = 0.99
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub RndPassthroughWorks()
    On Error GoTo TestFail

    With Fakes.Rnd
        .PassThrough = True
        Debug.Print Rnd()
        .Verify.Once
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub GetSettingFakeWorks()
    On Error GoTo TestFail

    With Fakes.GetSetting
        .Returns "Fakes work!"
        Dim retVal As String
        retVal = GetSetting("MyApp", "MySection", "MyKey")
        .Verify.Once
        .Verify.Parameter "appname", "MyApp"
        .Verify.Parameter "section", "MySection"
        .Verify.Parameter "key", "MyKey"
        Assert.IsTrue retVal = "Fakes work!"
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub GetSettingFakePassThroughWorks()
    On Error GoTo TestFail

    With Fakes.GetSetting
        .PassThrough = True
        GetSetting "MyApp", "MySection", "MyKey"
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
Public Sub GetAllSettingsFakeWorks()
    On Error GoTo TestFail

    With Fakes.GetAllSettings
        Dim fakeReturn As Variant
        ReDim fakeReturn(0 To 1, 0 To 1)
        fakeReturn(0, 0) = "Fake key 1"
        fakeReturn(0, 1) = "Fake setting 1"
        fakeReturn(1, 0) = "Fake key 2"
        fakeReturn(1, 1) = "Fake setting 2"
        .Returns fakeReturn
        
        Dim retVal As Variant
        retVal = GetAllSettings("MyApp", "MySection")
        .Verify.Once
        .Verify.Parameter "appname", "MyApp"
        .Verify.Parameter "section", "MySection"
        Assert.IsTrue retVal(0, 0) = fakeReturn(0, 0)
        Assert.IsTrue retVal(0, 1) = fakeReturn(0, 1)
        Assert.IsTrue retVal(1, 0) = fakeReturn(1, 0)
        Assert.IsTrue retVal(1, 1) = fakeReturn(1, 1)
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub GetAllSettingsFakePassThroughWorks()
    On Error GoTo TestFail

    With Fakes.GetAllSettings
        .PassThrough = True
        GetAllSettings "MyApp", "MySection"
        .Verify.Once
        .Verify.Parameter "appname", "MyApp"
        .Verify.Parameter "section", "MySection"
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub GetAttrFakeWorks()
    On Error GoTo TestFail

    With Fakes.GetAttr
        .Returns vbHidden + vbReadOnly
        Dim retVal As Integer
        retVal = GetAttr("C:\Test\dummy.txt")
        .Verify.Once
        .Verify.Parameter "pathname", "C:\Test\dummy.txt"
        Assert.IsTrue retVal = (vbHidden + vbReadOnly)
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub GetAttrFakePassThroughWorks()
    On Error GoTo TestFail

    With Fakes.GetAttr
        .PassThrough = True
        Dim retVal As Integer
        retVal = GetAttr("C:\Test\dummy.txt")
        .Verify.Once
        .Verify.Parameter "pathname", "C:\Test\dummy.txt"
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub FileLenFakeWorks()
    On Error GoTo TestFail

    With Fakes.FileLen
        .Returns 1234
        Dim retVal As Integer
        retVal = FileLen("C:\Test\dummy.txt")
        .Verify.Once
        .Verify.Parameter "pathname", "C:\Test\dummy.txt"
        Assert.IsTrue retVal = 1234
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub FileLenFakePassThroughWorks()
    On Error GoTo TestFail

    With Fakes.FileLen
        .PassThrough = True
        Dim retVal As Integer
        retVal = FileLen("C:\Test\dummy.txt")
        .Verify.Once
        .Verify.Parameter "pathname", "C:\Test\dummy.txt"
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub FileDateTimeFakeWorks()
    On Error GoTo TestFail

    With Fakes.FileDateTime
        .Returns DateSerial(2022, 11, 6) + TimeSerial(12, 0, 0)
        Dim retVal As Double
        retVal = FileDateTime("C:\Test\dummy.txt")
        .Verify.Once
        .Verify.Parameter "pathname", "C:\Test\dummy.txt"
        Assert.IsTrue retVal = DateSerial(2022, 11, 6) + TimeSerial(12, 0, 0)
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub FileDateTimeFakePassThroughWorks()
    On Error GoTo TestFail

    With Fakes.FileDateTime
        .PassThrough = True
        Dim retVal As Double
        retVal = FileDateTime("C:\Test\dummy.txt")
        .Verify.Once
        .Verify.Parameter "pathname", "C:\Test\dummy.txt"
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub IMEStatusFakeWorks()
    On Error GoTo TestFail

    With Fakes.IMEStatus
        .Returns vbIMEModeAlpha
        Dim retVal As Integer
        retVal = IMEStatus()
        .Verify.Once
        Assert.IsTrue retVal = vbIMEModeAlpha
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub IMEStatusFakePassThroughWorks()
    On Error GoTo TestFail

    With Fakes.IMEStatus
        .PassThrough = True
        Dim retVal As Integer
        retVal = IMEStatus()
        .Verify.Once
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub FreeFileFakeWorks()
    On Error GoTo TestFail

    With Fakes.FreeFile
        .Returns 300
        Dim retVal As Integer
        retVal = FreeFile(1)
        .Verify.Once
        .Verify.Parameter "rangenumber", 1
        Assert.IsTrue retVal = 300
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub FreeFileFakePassThroughWorks()
    On Error GoTo TestFail

    With Fakes.FreeFile
        .PassThrough = True
        Dim retVal As Integer
        retVal = FreeFile()
        .Verify.Once
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub DirFakeWorks()
    On Error GoTo TestFail

    With Fakes.Dir
        .Returns "File1.txt", 1
        .Returns "File2.txt", 2
        .Returns "", 3
        Dim retVal As String
        retVal = Dir("C:\Nothing", vbHidden)
        .Verify.Once
        .Verify.Parameter "pathname", "C:\Nothing"
        .Verify.Parameter "attributes", CInt(vbHidden)
        Assert.IsTrue retVal = "File1.txt"
        retVal = Dir()
        Assert.IsTrue retVal = "File2.txt"
        retVal = Dir()
        Assert.IsTrue retVal = ""
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub DirFakePassThroughWorks()
    On Error GoTo TestFail

    With Fakes.Dir
        .PassThrough = True
        Dim retVal As String
        retVal = Dir("C:\Test")
        .Verify.Once
        .Verify.Parameter "pathname", "C:\Test"
        .Verify.Parameter "attributes", 0
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestIssue4476()
    On Error GoTo TestFail
    
    'Arrange:
    Fakes.Now.PassThrough = True
    Fakes.Date.PassThrough = True
    Dim retVal As Variant
    
    'Act:
    retVal = Now
    retVal = Date '<== KA-BOOOM
    retVal = Now 'ensure fake reinstated
    
    'Assert:
    Fakes.Now.Verify.Exactly 2
    Fakes.Date.Verify.Once
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestIssue5944()
    On Error GoTo TestFail
        
    Fakes.InputBox.Returns 20
    Fakes.MsgBox.Returns 20
    
    Dim inputBoxReturnValue As String
    Dim msgBoxReturnValue As Integer
    
    inputBoxReturnValue = InputBox("Dummy")
    msgBoxReturnValue = MsgBox("Dummy")
    
    Fakes.MsgBox.Verify.Once
    Fakes.InputBox.Verify.Once
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

