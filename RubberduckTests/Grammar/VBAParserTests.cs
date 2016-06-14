﻿using Antlr4.Runtime;
using Antlr4.Runtime.Atn;
using Antlr4.Runtime.Tree;
using Antlr4.Runtime.Tree.Xpath;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;

namespace RubberduckTests.Grammar
{
    [TestClass]
    public class VBAParserTests
    {
        [TestMethod]
        public void TestParsesEmptyForm()
        {
            var code = @"
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form1 
   Caption         =   ""Form1""
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   ""Form1.frx"":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = ""Form1""
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//attributeStmt");
        }

        [TestMethod]
        public void TestAttributeFirstLine()
        {
            string code = @"
Attribute VB_Name = ""Form1""
VERSION 5.00";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//attributeStmt");
        }

        [TestMethod]
        public void TestAttributeAfterModuleHeader()
        {
            string code = @"
VERSION 5.00
Attribute VB_Name = ""Form1""
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form1 
   Caption         =   ""Form1""
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   ""Form1.frx"":0000
   StartUpPosition =   1  'CenterOwner
End
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//attributeStmt");
        }

        [TestMethod]
        public void TestAttributeAfterModuleConfig()
        {
            string code = @"
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form1 
   Caption         =   ""Form1""
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   ""Form1.frx"":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = ""Form1""
Private this As TProgressIndicator
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//attributeStmt");
        }

        [TestMethod]
        public void TestAttributeInsideModuleDeclarations()
        {
            string code = @"
Public WithEvents colCBars As Office.CommandBars
Attribute colCBars.VB_VarHelpID = -1
Public WithEvents colCBars2 As Office.CommandBars
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//attributeStmt");
        }

        [TestMethod]
        public void TestAttributeAfterModuleDeclarations()
        {
            string code = @"
Private this As TProgressIndicator
Attribute VB_Name = ""Form1""
Public Sub Test()
    Attribute VB_Name = ""Form1""
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//attributeStmt");
        }

        [TestMethod]
        public void TestAttributeInsideProcedure()
        {
            string code = @"
Public Sub Test()
    Attribute VB_Name = ""Form1""
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//attributeStmt");
        }

        [TestMethod]
        public void TestAttributeEndOfFile()
        {
            string code = @"
Public Sub Test()
End Sub
Attribute VB_Name = ""Form1""
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//attributeStmt");
        }

        [TestMethod]
        public void TestAttributeNameIsMemberAccessExpr()
        {
            string code = @"
Attribute view.VB_VarHelpID = -1
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//attributeStmt");
        }

        [TestMethod]
        public void TestTrivialCase()
        {
            string code = @":";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//module");
        }

        [TestMethod]
        public void TestEmptyModule()
        {
            string code = @"
_

   _

           _

";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//module");
        }

        [TestMethod]
        public void TestModuleHeader()
        {
            string code = @"VERSION 1.0 CLASS";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleHeader");
        }

        [TestMethod]
        public void TestDefDirectiveSingleLetter()
        {
            string code = @"DefBool B: DefByte Y: DefInt I: DefLng L: DefLngLng N: DefLngPtr P: DefCur C: DefSng G: DefDbl D: DefDate T: DefStr E: DefObj O: DefVar V";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//defDirective", matches => matches.Count == 13);
        }

        [TestMethod]
        public void TestDefDirectiveSameDefDirectiveMultipleLetterSpec()
        {
            string code = @"DefBool B, C, D";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//singleLetter", matches => matches.Count == 3);
        }

        [TestMethod]
        public void TestDefDirectiveLetterRange()
        {
            string code = @"DefBool B-C: DefByte Y-X: DefInt I-J: DefLng L-M: DefLngLng N-O: DefLngPtr P-Q: DefCur C-D: DefSng G-H: DefDbl D-E: DefDate T-U: DefStr E-F: DefObj O-P: DefVar V-W";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//letterRange", matches => matches.Count == 13);
        }

        [TestMethod]
        public void TestDefDirectiveUniversalLetterRange()
        {
            string code = @"DefBool A - Z";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//universalLetterRange");
        }

        [TestMethod]
        public void TestModuleOption()
        {
            string code = @"
Option Explicit

Sub DoSomething()
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleOption");
        }

        [TestMethod]
        public void TestModuleOption_Indented()
        {
            string code = @"
    Option Explicit

    Sub DoSomething()
    End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleOption");
        }

        [TestMethod]
        public void TestModuleConfig()
        {
            string code = @"
BEGIN
  MultiUse = -1  'True
END";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleConfigElement");
        }

        [TestMethod]
        public void TestEmptyComment()
        {
            string code = @"'";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//comment");
        }

        [TestMethod]
        public void TestEmptyRemComment()
        {
            string code = @"Rem";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//remComment");
        }

        [TestMethod]
        public void TestOneCharRemComment()
        {
            string code = @"Rem a";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//remComment");
        }

        [TestMethod]
        public void TestCommentThatLooksLikeAnnotation()
        {
            string code = @"'@param foo: the value of something";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//comment");
        }

        [TestMethod]
        public void TestForeignIdentifier()
        {
            string code = @"
Sub FooFoo()
  [Sheet2!A2]
  [[Book2]Sheet1!A1]
  [Book2!NamedRange]
  [""Hello World!""]
  [""!""]
  [""[]""]
  []
  a = [A1] + [A2]
End Sub";
            var parseResult = Parse(code);
            // foreign names + 1 for the subroutine's name.
            AssertTree(parseResult.Item1, parseResult.Item2, "//identifier", matches => matches.Count == 11);
        }

        [TestMethod]
        public void TestOneCharComment()
        {
            string code = @"'a";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//comment");
        }

        [TestMethod]
        public void TestEndEnumMultipleWhiteSpace()
        {
            string code = @"
Enum Test
    anything
End               Enum";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//enumerationStmt");
        }

        [TestMethod]
        public void TestEndTypeMultipleWhiteSpace()
        {
            string code = @"
Type Test
    anything As Integer
End             Type";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//udtDeclaration");
        }

        [TestMethod]
        public void TestEndFunctionLineContinuation()
        {
            string code = @"
Function Test()

End _
Function";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//functionStmt");
        }

        [TestMethod]
        public void TestExitFunctionLineContinuation()
        {
            string code = @"
Public Function Test()
    Exit _
    Function
End Function";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//functionStmt");
        }

        [TestMethod]
        public void TestEndSubroutineLineContinuation()
        {
            string code = @"
Sub Test()

End _
Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//subStmt");
        }

        [TestMethod]
        public void TestExitSubroutineLineContinuation()
        {
            string code = @"
Sub Test()
    Exit _
    Sub
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//subStmt");
        }

        [TestMethod]
        public void TestPropertyGetLineContinuation()
        {
            string code = @"
Property _
Get Test()
End Property";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//propertyGetStmt");
        }

        [TestMethod]
        public void TestPropertyLetLineContinuation()
        {
            string code = @"
Property _
Let Test(anything As Integer)
End Property";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//propertyLetStmt");
        }

        [TestMethod]
        public void TestPropertySetLineContinuation()
        {
            string code = @"
Property _
Set Test(anything As Application)
End Property";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//propertySetStmt");
        }

        [TestMethod]
        public void TestEndPropertyLineContinuation()
        {
            string code = @"
Property Get Test()

End _
Property";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//propertyGetStmt");
        }

        [TestMethod]
        public void TestExitPropertyLineContinuation()
        {
            string code = @"
Public Property Get Test()
    Exit _
    Property
End Property";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//propertyGetStmt");
        }

        [TestMethod]
        public void TestEndIfLineContinuation()
        {
            string code = @"
Function Test()
    If 1 = 1 Then
    End _
    If
End Function";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//ifStmt");
        }

        [TestMethod]
        public void TestEndSelectLineContinuation()
        {
            string code = @"
Property Get Test()
    Select Case 1 = 2
    End _
    Select
End Property";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//selectCaseStmt");
        }

        [TestMethod]
        public void TestEndWithContinuation()
        {
            string code = @"
Sub Test()
  With Application
  End _
  With
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//withStmt");
        }

        [TestMethod]
        public void TestExitDoContinuation()
        {
            string code = @"
Sub Test()
    Do While True
        Exit _
        Do
    Loop
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//doLoopStmt");
        }

        [TestMethod]
        public void TestExitForContinuation()
        {
            string code = @"
Sub Test()
    For i = 1 To 2
        Exit _
        For
    Next i
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//forNextStmt");
        }

        [TestMethod]
        public void TestLineInputLineContinuation()
        {
            string code = @"
Sub Test()
    Line _
    Input #1, TextLine
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//lineInputStmt");
        }

        [TestMethod]
        public void TestReadWriteKeywordLineContinuation()
        {
            string code = @"
Sub Test()
    Open ""TESTFILE"" For Random Access Read _
    Write As #1
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//openStmt");
        }

        [TestMethod]
        public void TestLockReadKeywordLineContinuation()
        {
            string code = @"
Sub Test()
    Open ""TESTFILE"" For Random Lock _
    Read As #1
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//openStmt");
        }

        [TestMethod]
        public void TestLockWriteKeywordLineContinuation()
        {
            string code = @"
Sub Test()
    Open ""TESTFILE"" For Random Lock _
    Write As #1
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//openStmt");
        }

        [TestMethod]
        public void TestLockReadWriteKeywordLineContinuation()
        {
            string code = @"
Sub Test()
    Open ""TESTFILE"" For Random Lock _
    Read _
    Write As #1
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//openStmt");
        }

        [TestMethod]
        public void TestOnErrorLineContinuation()
        {
            string code = @"
Sub Test()
On _
Error GoTo a
a:
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//onErrorStmt");
        }

        [TestMethod]
        public void TestOnLocalErrorLineContinuation()
        {
            string code = @"
Sub Test()
On _
Local _
Error GoTo a
a:
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//onErrorStmt");
        }

        [TestMethod]
        public void TestOptionBaseLineContinuation()
        {
            string code = @"
Option _
Base _
1";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleOption");
        }

        [TestMethod]
        public void TestOptionExplicitLineContinuation()
        {
            string code = @"
Option _
Explicit";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleOption");
        }

        [TestMethod]
        public void TestOptionCompareLineContinuation()
        {
            string code = @"
Option _
Compare _
Text";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleOption");
        }

        [TestMethod]
        public void TestOptionPrivateModuleLineContinuation()
        {
            string code = @"
Option _
Private _
Module";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleOption");
        }

        [TestMethod]
        public void TestDictionaryAccessExprLineContinuation()
        {
            string code = @"
Public Sub Test()
    Set result = myObj _
    ! _
    call
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//lExpression");
        }

        [TestMethod]
        public void TestWithDictionaryAccessExprLineContinuation()
        {
            string code = @"
Public Sub Test()
    With Application
        ! _ 
  Activate
    End With
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//lExpression");
        }

        [TestMethod]
        public void TestLetStmtLineContinuation()
        {
            string code = @"
Public Sub Test()
    x = ( _
            1 / _
            1 _
        ) * 1
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//letStmt");
        }

        [TestMethod]
        public void TestMemberAccessExprLineContinuation()
        {
            string code = @"
Public Sub Test()
    Debug.Print Foo.Bar _
                   . _
                    FooBar.Baz
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//lExpression");
        }

        [TestMethod]
        public void TestWithMemberAccessExprLineContinuation()
        {
            string code = @"
Public Sub Test()
    With Application
        . _
    Activate
    End With
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//lExpression");
        }

        [TestMethod]
        public void TestCallStmtLineContinuation()
        {
            string code = @"
Sub Test()
	fun(1) _
	.fun(2) _
	.fun(3)
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//lExpression");
        }

        [TestMethod]
        public void TestDeclareLineContinuation()
        {
            string code = @"
Private Declare Function ABC Lib ""shell32.dll"" Alias _
""ShellExecuteA""(ByVal a As Long, ByVal b As String, _
ByVal c As String, ByVal d As String, ByVal e As String, ByVal f As Long) As Long";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//declareStmt");
        }

        [TestMethod]
        public void TestEraseStmt()
        {
            string code = @"
Public Sub EraseTwoArrays()
Erase someArray(), someOtherArray()
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//eraseStmt");
        }

        [TestMethod]
        public void TestFixedLengthString()
        {
            string code = @"
Sub Test()
    Dim someString As String * 255
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//fieldLength");
        }

        [TestMethod]
        public void TestDoLoopStatement()
        {
            string code = @"
Sub Test()
    Do
    Loop Until var > 10
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//doLoopStmt");
        }

        [TestMethod]
        public void TestForNextStatement()
        {
            string code = @"
Sub Test()
    For n = 1 To 10
    Next n%
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//forNextStmt");
        }

        [TestMethod]
        public void TestLineLabelStatement()
        {
            string code = @"
Sub Test()
    a:
    10:
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//statementLabelDefinition", matches => matches.Count == 2);
        }

        [TestMethod]
        public void TestAnnotations()
        {
            string code = @"
'@Folder a @Folder b
Sub Test()
    ' Test Comment
    Dim someString As String * 255 '@Folder c @Folder d
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//annotation", matches => matches.Count == 4);
        }

        [TestMethod]
        public void TestEmptyAnnotationsWithParentheses()
        {
            string code = @"
'@NoIndent()
Sub Test()
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//annotation");
        }

        [TestMethod]
        public void GivenIfElseBlock_ParsesElseBlockAsElseStatement()
        {
            var code = @"
Private Sub DoSomething()
    If Not True Then
        Debug.Print False
    Else
        Debug.Print True
    End If
End Sub
";
            var parser = Parse(code);
            AssertTree(parser.Item1, parser.Item2, "//elseBlock", matches => matches.Count == 1);
        }

        [TestMethod]
        public void TestIfStmtSameLineElse()
        {
            string code = @"
Sub Test()
    If True Then
    ElseIf False Then Debug.Print 5
    Else
    End If
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//elseIfBlock");
        }

        [TestMethod]
        public void TestSingleLineIfEmptyThenEmptyElse()
        {
            string code = @"
Sub Test()
    If False Then Else
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//singleLineIfStmt");
        }

        [TestMethod]
        public void TestSingleLineIfEmptyThenEndOfStatement()
        {
            string code = @"
Sub Test()
    If False Then: Else
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//singleLineIfStmt");
        }

        [TestMethod]
        public void TestSingleLineIfMultipleThenNoElse()
        {
            string code = @"
Sub Test()
      If False Then MsgBox False: MsgBox False Else
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//singleLineIfStmt");
        }

        [TestMethod]
        public void TestSingleLineIfMultipleThenMultipleElse()
        {
            string code = @"
Sub Test()
      If False Then MsgBox False: MsgBox False Else MsgBox False: MsgBox False
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//singleLineIfStmt");
        }

        [TestMethod]
        public void TestSingleLineIfEmptyThen()
        {
            string code = @"
Sub Test()
      If False Then Else MsgBox True
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//singleLineIfStmt");
        }

        [TestMethod]
        public void TestSingleLineIfSingleThenEmptyElse()
        {
            string code = @"
Sub Test()
      If False Then MsgBox True Else
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//singleLineIfStmt");
        }

        [TestMethod]
        public void TestSingleLineIfImplicitGoTo()
        {
            string code = @"
Sub Test()
      ' This actually means: If True Then GoTo 5 Else GoTo 10
      If True Then 5 Else 10
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//lineNumberLabel", matches => matches.Count == 2);
        }

        [TestMethod]
        public void TestSingleLineIfDoLoop()
        {
            string code = @"
Sub Test()
      If True Then Do: Loop Else
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//doLoopStmt");
        }

        [TestMethod]
        public void TestSingleLineIfWendLoop()
        {
            string code = @"
Sub Test()
      If True Then While True: Beep: Wend Else
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//whileWendStmt");
        }

        [TestMethod]
        public void TestSingleLineIfRealWorldExample1()
        {
            string code = @"
Sub Test()
      On Local Error Resume Next: If Not Empty Is Nothing Then Do While Null: ReDim i(True To False) As Currency: Loop: Else Debug.Assert CCur(CLng(CInt(CBool(False Imp True Xor False Eqv True)))): Stop: On Local Error GoTo 0
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//singleLineIfStmt");
        }

        [TestMethod]
        public void TestSingleLineIfRealWorldExample2()
        {
            string code = @"
Sub Test()
    With Application
        If bUpdate Then .Calculation = xlCalculationAutomatic: .Cursor = xlDefault Else .Calculation = xlCalculationManual: .Cursor = xlWait: .EnableEvents = bUpdate: .ScreenUpdating = bUpdate: .DisplayAlerts = bUpdate: .CutCopyMode = False: .StatusBar = False
    End With
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//singleLineIfStmt");
        }

        [TestMethod]
        public void TestSingleLineIfRealWorldExample3()
        {
            string code = @"
Sub Test()
    If Not oP_Window Is Nothing Then If Not oP_Window.Visible Then Unload oP_Window: Set oP_Window = Nothing
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//singleLineIfStmt", matches => matches.Count == 2);
        }

        [TestMethod]
        public void TestSingleLineIfRealWorldExample4()
        {
            string code = @"
Sub Test()
    If Err Then Set oP_Window = Nothing: TurnOff Else If oP_Window Is Nothing Then TurnOn
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//singleLineIfStmt", matches => matches.Count == 2);
        }

        [TestMethod]
        public void TestEndStmt()
        {
            string code = @"
Sub Test()
    End
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//endStmt");
        }

        [TestMethod]
        public void TestRedimStmtArray()
        {
            string code = @"
Sub Test()
    ReDim strArray(1)
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//expression");
        }

        [TestMethod]
        public void TestRedimStmtLowerBoundsArgument()
        {
            string code = @"
Sub Test()
    ReDim strArray(1 To 10)
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//lowerBoundArgumentExpression");
        }

        [TestMethod]
        public void TestRedimStmtUpperBoundsArgument()
        {
            string code = @"
Sub Test()
    ReDim strArray(1 To 10)
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//upperBoundArgumentExpression");
        }

        [TestMethod]
        public void TestRedimStmtNormalArgument()
        {
            string code = @"
Sub Test()
    ReDim strArray(1 To 10)
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//argumentExpression");
        }

        [TestMethod]
        public void TestStringFunction()
        {
            string code = @"
Sub Test()
    a = String(5, ""a"")
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//lExpression");
        }

        [TestMethod]
        public void TestArrayWithTypeSuffix()
        {
            string code = @"
Sub Test()
    Dim a!()
    Dim a@()
    Dim a#()
    Dim a$()
    Dim a%()
    Dim a^()
    Dim a&()
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//typeHint", matches => matches.Count == 7);
        }

        [TestMethod]
        public void TestOpenStmt()
        {
            string code = @"
Sub Test()
    Open ""TESTFILE"" For Binary Access Read Lock Read As #1 Len = 2
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//openStmt");
        }

        [TestMethod]
        public void TestResetStmt()
        {
            string code = @"
Sub Test()
    Reset
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//resetStmt");
        }

        [TestMethod]
        public void TestCloseStmt()
        {
            string code = @"
Sub Test()
    Close #1, 2, 3
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//closeStmt");
        }

        [TestMethod]
        public void TestSeekStmt()
        {
            string code = @"
Sub Test()
    Seek #1, 2
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//seekStmt");
        }

        [TestMethod]
        public void TestSeekFunction()
        {
            // Tests whether SEEK, which is actually a special keyword, can also be used in a "function call context".
            string code = @"
Sub Test()
    anything = Seek(50)
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//lExpression");
        }

        [TestMethod]
        public void TestLockStmt()
        {
            string code = @"
Sub Test()
    Lock #1, 2
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//lockStmt");
        }

        [TestMethod]
        public void TestUnlockStmt()
        {
            string code = @"
Sub Test()
    Unlock #1, 2
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//unlockStmt");
        }

        [TestMethod]
        public void TestLineInputStmt()
        {
            string code = @"
Sub Test()
    Line Input #2, ""ABC""
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//lineInputStmt");
        }

        [TestMethod]
        public void TestWidthStmt()
        {
            string code = @"
Sub Test()
    Width #2, 5
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//widthStmt");
        }

        [TestMethod]
        public void TestPrintStmt()
        {
            string code = @"
Sub Test()
    Print #2, Spc(5) ;
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//printStmt");
        }

        [TestMethod]
        public void TestDebugPrintStmtNoArguments()
        {
            string code = @"
Sub Test()
    Debug.Print
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//debugPrintStmt");
        }

        [TestMethod]
        public void TestDebugPrintStmtNormalArgumentSyntax()
        {
            string code = @"
Sub Test()
    Debug.Print ""Anything""
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//debugPrintStmt/outputList");
        }

        [TestMethod]
        public void TestDebugPrintStmtOutputItemSemicolon()
        {
            string code = @"
Sub Test()
    Debug.Print 1;
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//debugPrintStmt/outputList");
        }

        [TestMethod]
        public void TestDebugPrintStmtOutputItemComma()
        {
            string code = @"
Sub Test()
    Debug.Print 1,
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//debugPrintStmt/outputList");
        }

        [TestMethod]
        public void TestDebugPrintRealWorldExample1()
        {
            string code = @"
Sub Test()
    For Each fld In tdf.Fields
        Debug.Print fld.Name,
        Debug.Print FieldTypeName(fld),
        Debug.Print fld.Size,
        Debug.Print GetDescrip(fld)
    Next
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//debugPrintStmt", matches => matches.Count == 4);
        }

        [TestMethod]
        public void TestDebugPrintRealWorldExample2()
        {
            string code = @"
Sub Test()
    If Not pFault Then
        Debug.Print ""FirstO: "" & vbCr & ans(0) & vbCr
        Debug.Print ""SecondO:""; ans(1)
    End If
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//debugPrintStmt", matches => matches.Count == 2);
        }

        [TestMethod]
        public void TestDebugPrintRealWorldExample3()
        {
            string code = @"
Sub Test()
    For i = LBound(sortedArray) To UBound(sortedArray)
        Debug.Print sortedArray(i) & "":"";
    Next i
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//debugPrintStmt", matches => matches.Count == 1);
        }

        [TestMethod]
        public void TestWriteStmt()
        {
            string code = @"
Sub Test()
    Write #1, ""ABC"", 234
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//writeStmt");
        }

        [TestMethod]
        public void TestInputStmt()
        {
            string code = @"
Sub Test()
    Input #1, ""ABC""
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//inputStmt");
        }

        [TestMethod]
        public void TestInputFunction()
        {
            string code = @"
Sub Test()
    s = Input(LOF(file1), #file1)
    s = Input$(LOF(file1), #file1)
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//lExpression");
        }

        [TestMethod]
        public void TestInputBFunction()
        {
            string code = @"
Sub Test()
    s = InputB(LOF(file1), #file1)
    s = InputB$(LOF(file1), #file1)
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//lExpression");
        }

        [TestMethod]
        public void TestCircleSpecialForm()
        {
            string code = @"
Sub Test()
    Me.Circle Step(1, 2), 3, 4, 5, 6, 7
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//circleSpecialForm");
        }

        [TestMethod]
        public void TestScaleSpecialForm()
        {
            string code = @"
Sub Test()
    Scale (1, 2)-(3, 4)
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//scaleSpecialForm");
        }

        [TestMethod]
        public void TestPtrSafeAsSub()
        {
            string code = @"
Private Sub PtrSafe()
    Debug.Print 42
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//subStmt");
        }

        [TestMethod]
        public void TestFunction_Indented()
        {
            string code = @"
    Private Function Foo() As Boolean
        Foo = True
    End Function";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//functionStmt");
        }

        [TestMethod]
        public void TestSub_Indented()
        {
            string code = @"
    Private Sub Foo()
    End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//subStmt");
        }

        [TestMethod]
        public void TestSub_InconsistentlyIndented()
        {
            string code = @"
    Private Sub Foo()
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//subStmt");
        }

        [TestMethod]
        public void TestPtrSafeAsVariable()
        {
            string code = @"
Private Sub Foo()
    Dim PtrSafe As Integer
    PtrSafe = 42
    Debug.Print PtrSafe
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//variableStmt");
        }

        [TestMethod]
        public void TestLiteralExpressionResolvesCorrectly()
        {
            string code = @"
Private Sub Foo()
    a = True
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//literalExpression");
        }

        [TestMethod]
        public void TestUdtReservedKeywords()
        {
            string code = @"
Private Type Foo
    If As Integer
    Select As Integer
    Split As String
    For As Integer
    Dim As Integer
    Then As Integer
    UBound As Variant
    To As Integer
    Or As Integer
    Case As Integer
    Type As Integer
    Enum As Integer
    End As Integer
End Type
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//udtMember", matches => matches.Count == 13);
        }

        private Tuple<VBAParser, ParserRuleContext> Parse(string code)
        {
            var stream = new AntlrInputStream(code);
            var lexer = new VBALexer(stream);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBAParser(tokens);
            // Don't remove this line otherwise we won't get notified of parser failures.
            parser.AddErrorListener(new ExceptionErrorListener());
            // If SLL fails we want to get notified ASAP so we can fix it, that's why we don't retry using LL.
            parser.Interpreter.PredictionMode = PredictionMode.Sll;
            var tree = parser.startRule();
            return Tuple.Create<VBAParser, ParserRuleContext>(parser, tree);
        }

        private void AssertTree(VBAParser parser, ParserRuleContext root, string xpath)
        {
            AssertTree(parser, root, xpath, matches => matches.Count >= 1);
        }

        private void AssertTree(VBAParser parser, ParserRuleContext root, string xpath, Predicate<ICollection<IParseTree>> assertion)
        {
            var matches = new XPath(parser, xpath).Evaluate(root);
            var actual = matches.Count;
            Assert.IsTrue(assertion(matches), string.Format("{0} matches found.", actual));
        }
    }
}