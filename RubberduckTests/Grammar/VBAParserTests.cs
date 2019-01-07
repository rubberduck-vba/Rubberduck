using Antlr4.Runtime;
using Antlr4.Runtime.Atn;
using Antlr4.Runtime.Tree;
using Antlr4.Runtime.Tree.Xpath;
using NUnit.Framework;
using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace RubberduckTests.Grammar
{
    [TestFixture]
    public class VBAParserTests
    {
        [Category("Parser")]
        [Category("LineLabels")]
        [Test]
        public void DoEventsKeywordDoesNotParseAsLineLabel()
        {
            var code = @"
Sub DoSomething()
DoEvents: MsgBox
End Sub
";
            var result = Parse(code);
            AssertTree(result.Item1, result.Item2, "//statementLabelDefinition", e => e.Count == 0);
            AssertTree(result.Item1, result.Item2, "//subStmt");
        }

        [Category("Parser")]
        [Category("LineLabels")]
        [Test]
        public void EndKeywordDoesNotParseAsLineLabel()
        {
            var code = @"
Sub DoSomething()
End: MsgBox
End Sub
";
            var result = Parse(code);
            AssertTree(result.Item1, result.Item2, "//statementLabelDefinition", e => e.Count == 0);
            AssertTree(result.Item1, result.Item2, "//subStmt");
        }

        [Category("Parser")]
        [Category("LineLabels")]
        [Test]
        public void CloseKeywordDoesNotParseAsLineLabel()
        {
            var code = @"
Sub DoSomething()
Close: MsgBox
End Sub
";
            var result = Parse(code);
            AssertTree(result.Item1, result.Item2, "//statementLabelDefinition", e => e.Count == 0);
            AssertTree(result.Item1, result.Item2, "//subStmt");
        }

        [Category("Parser")]
        [Category("LineLabels")]
        [Test]
        public void DoKeywordDoesNotParseAsLineLabel()
        {
            var code = @"
Sub DoSomething()
Do: MsgBox : Loop
End Sub
";
            var result = Parse(code);
            AssertTree(result.Item1, result.Item2, "//statementLabelDefinition", e => e.Count == 0);
            AssertTree(result.Item1, result.Item2, "//doLoopStmt");
            AssertTree(result.Item1, result.Item2, "//subStmt");
        }

        [Category("Parser")]
        [Category("LineLabels")]
        [Test]
        public void ElseKeywordDoesNotParseAsLineLabel()
        {
            var code = @"
Sub DoSomething()
If True Then
Else: MsgBox
End If
End Sub
";
            var result = Parse(code);
            AssertTree(result.Item1, result.Item2, "//statementLabelDefinition", e => e.Count == 0);
            AssertTree(result.Item1, result.Item2, "//ifStmt");
            AssertTree(result.Item1, result.Item2, "//subStmt");
        }

        [Category("Parser")]
        [Category("LineLabels")]
        [Test]
        public void LoopKeywordDoesNotParseAsLineLabel()
        {
            var code = @"
Sub DoSomething()
Do Until False
Loop: MsgBox
End Sub
";
            var result = Parse(code);
            AssertTree(result.Item1, result.Item2, "//statementLabelDefinition", e => e.Count == 0);
            AssertTree(result.Item1, result.Item2, "//doLoopStmt");
            AssertTree(result.Item1, result.Item2, "//subStmt");
        }

        [Category("Parser")]
        [Category("LineLabels")]
        [Test]
        public void NextKeywordDoesNotParseAsLineLabel()
        {
            var code = @"
Sub DoSomething()
For i = 1 To 10
Next: MsgBox
End Sub
";
            var result = Parse(code);
            AssertTree(result.Item1, result.Item2, "//statementLabelDefinition", e => e.Count == 0);
            AssertTree(result.Item1, result.Item2, "//forNextStmt");
            AssertTree(result.Item1, result.Item2, "//subStmt");
        }

        [Category("Parser")]
        [Category("LineLabels")]
        [Test]
        public void RandomizeKeywordDoesNotParseAsLineLabel()
        {
            var code = @"
Sub DoSomething()
Randomize: MsgBox
End Sub
";
            var result = Parse(code);
            AssertTree(result.Item1, result.Item2, "//statementLabelDefinition", e => e.Count == 0);
            AssertTree(result.Item1, result.Item2, "//subStmt");
        }

        [Category("Parser")]
        [Category("LineLabels")]
        [Test]
        public void RemKeywordDoesNotParseAsLineLabel()
        {
            var code = @"
Sub DoSomething()
Rem: MsgBox
End Sub
";
            var result = Parse(code);
            AssertTree(result.Item1, result.Item2, "//statementLabelDefinition", e => e.Count == 0);
            AssertTree(result.Item1, result.Item2, "//remComment");
            AssertTree(result.Item1, result.Item2, "//subStmt");
        }

        [Category("Parser")]
        [Category("LineLabels")]
        [Test]
        public void ResumeKeywordDoesNotParseAsLineLabel()
        {
            var code = @"
Sub DoSomething()
Resume: MsgBox
End Sub
";
            var result = Parse(code);
            AssertTree(result.Item1, result.Item2, "//statementLabelDefinition", e => e.Count == 0);
            AssertTree(result.Item1, result.Item2, "//resumeStmt");
            AssertTree(result.Item1, result.Item2, "//subStmt");
        }

        [Category("Parser")]
        [Category("LineLabels")]
        [Test]
        public void ReturnKeywordDoesNotParseAsLineLabel()
        {
            var code = @"
Sub DoSomething()
Return: MsgBox
End Sub
";
            var result = Parse(code);
            AssertTree(result.Item1, result.Item2, "//statementLabelDefinition", e => e.Count == 0);
            AssertTree(result.Item1, result.Item2, "//returnStmt");
            AssertTree(result.Item1, result.Item2, "//subStmt");
        }

        [Category("Parser")]
        [Category("LineLabels")]
        [Test]
        public void StopKeywordDoesNotParseAsLineLabel()
        {
            var code = @"
Sub DoSomething()
Stop: MsgBox
End Sub
";
            var result = Parse(code);
            AssertTree(result.Item1, result.Item2, "//statementLabelDefinition", e => e.Count == 0);
            AssertTree(result.Item1, result.Item2, "//stopStmt");
            AssertTree(result.Item1, result.Item2, "//subStmt");
        }

        [Category("Parser")]
        [Category("LineLabels")]
        [Test]
        public void WendKeywordDoesNotParseAsLineLabel()
        {
            var code = @"
Sub DoSomething()
While True
Wend: MsgBox
End Sub
";
            var result = Parse(code);
            AssertTree(result.Item1, result.Item2, "//statementLabelDefinition", e => e.Count == 0);
            AssertTree(result.Item1, result.Item2, "//whileWendStmt");
            AssertTree(result.Item1, result.Item2, "//subStmt");
        }

        [Category("Parser")]
        [Test]
        public void ParsesWithLineNumbers_EndSub()
        {
            var code = @"
Sub DoSomething()
10 End Sub
";
            var result = Parse(code);
            AssertTree(result.Item1, result.Item2, "//statementLabelDefinition");
            AssertTree(result.Item1, result.Item2, "//subStmt");
        }

        [Category("Parser")]
        [Test]
        public void ParsesWithLineNumbers_EndFunction()
        {
            var code = @"
Function DoSomething()
10 End Function
";
            var result = Parse(code);
            AssertTree(result.Item1, result.Item2, "//statementLabelDefinition");
            AssertTree(result.Item1, result.Item2, "//functionStmt");
        }

        [Category("Parser")]
        [Test]
        public void ParsesWithLineNumbers_EndProperty()
        {
            var code = @"
Property Get DoSomething()
10 End Property
";
            var result = Parse(code);
            AssertTree(result.Item1, result.Item2, "//statementLabelDefinition");
            AssertTree(result.Item1, result.Item2, "//propertyGetStmt");
        }

        [Category("Parser")]
        [Test]
        public void ParsesWithLineNumbers_IfStmt()
        {
            var code = @"
Sub DoSomething()
10 If True Then Debug.Print 42
End Sub
";
            var result = Parse(code);
            AssertTree(result.Item1, result.Item2, "//statementLabelDefinition");
            AssertTree(result.Item1, result.Item2, "//singleLineIfStmt");
        }

        [Category("Parser")]
        [Test]
        public void ParsesWithLineNumbers_ElseStmt()
        {
            var code = @"
Sub DoSomething()
10 If True Then
11     Debug.Print 42
20 Else
21     Debug.Print 42
30 End If
End Sub
";
            var result = Parse(code);
            AssertTree(result.Item1, result.Item2, "//statementLabelDefinition", matches => matches.Count == 5);
            AssertTree(result.Item1, result.Item2, "//elseBlock");
        }

        [Category("Parser")]
        [Test]
        public void ParsesWithLineNumbers_SelectCaseStmt()
        {
            var code = @"
Sub DoSomething()
10 Select Case False
20 Case True
21     Debug.Print 42
30 Case False
31     Debug.Print 42
40 End Select
End Sub
";
            var result = Parse(code);
            AssertTree(result.Item1, result.Item2, "//statementLabelDefinition", matches => matches.Count == 6);
            AssertTree(result.Item1, result.Item2, "//caseClause", matches => matches.Count == 2);
        }

        [Category("Parser")]
        [Test]
        public void ParsesWithLineNumbers_ForNextLoop()
        {
            var code = @"
Sub DoSomething()
10 For i = 1 To 10
20     Debug.Print 42
30 Next
End Sub
";
            var result = Parse(code);
            AssertTree(result.Item1, result.Item2, "//statementLabelDefinition", matches => matches.Count == 3);
            AssertTree(result.Item1, result.Item2, "//forNextStmt");
        }

        [Category("Parser")]
        [Test]
        public void ParsesWithLineNumbers_ForEachLoop()
        {
            var code = @"
Sub DoSomething()
10 For Each foo In bar
20     Debug.Print 42
30 Next
End Sub
";
            var result = Parse(code);
            AssertTree(result.Item1, result.Item2, "//statementLabelDefinition", matches => matches.Count == 3);
            AssertTree(result.Item1, result.Item2, "//forEachStmt");
        }

        [Category("Parser")]
        [Test]
        public void ParsesWithLineNumbers_DoLoop()
        {
            var code = @"
Sub DoSomething()
10 Do
20     Debug.Print 42
30 Loop While False
End Sub
";
            var result = Parse(code);
            AssertTree(result.Item1, result.Item2, "//statementLabelDefinition", matches => matches.Count == 3);
            AssertTree(result.Item1, result.Item2, "//doLoopStmt");
        }

        [Category("Parser")]
        [Test]
        public void ParsesWithLineNumbers_WithBlock()
        {
            var code = @"
Sub DoSomething()
10 With New Collection
20     Debug.Print 42
30 End With
End Sub
";
            var result = Parse(code, PredictionMode.Sll);
            AssertTree(result.Item1, result.Item2, "//statementLabelDefinition", matches => matches.Count == 3);
            AssertTree(result.Item1, result.Item2, "//withStmt");
        }

        [Category("Parser")]
        [Test]
        public void ParsesWithLineNumbers_WhileLoop()
        {
            var code = @"
Sub DoSomething()
10 While False
20     Debug.Print 42
30 Wend
End Sub
";
            var result = Parse(code);
            AssertTree(result.Item1, result.Item2, "//statementLabelDefinition", matches => matches.Count == 3);
            AssertTree(result.Item1, result.Item2, "//whileWendStmt");
        }

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
        public void TestAttributeFirstLine()
        {
            string code = @"
Attribute VB_Name = ""Form1""
VERSION 5.00";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//attributeStmt");
        }

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
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


        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
        public void TestAttributeNameIsMemberAccessExpr()
        {
            string code = @"
Attribute view.VB_VarHelpID = -1
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//attributeStmt");
        }

        [Category("Parser")]
        [Test]
        public void TestTrivialCase()
        {
            string code = @":";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//module");
        }

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
        public void TestModuleHeader()
        {
            string code = @"VERSION 1.0 CLASS";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleHeader");
        }

        [Category("Parser")]
        [Test]
        public void TestDefDirectiveSingleLetter()
        {
            string code = @"DefBool B: DefByte Y: DefInt I: DefLng L: DefLngLng N: DefLngPtr P: DefCur C: DefSng G: DefDbl D: DefDate T: DefStr E: DefObj O: DefVar V";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//defDirective", matches => matches.Count == 13);
        }

        [Category("Parser")]
        [Test]
        public void TestDefDirectiveSameDefDirectiveMultipleLetterSpec()
        {
            string code = @"DefBool B, C, D";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//singleLetter", matches => matches.Count == 3);
        }

        [Category("Parser")]
        [Test]
        public void TestDefDirectiveLetterRange()
        {
            string code = @"DefBool A-C: DefByte Y-X: DefInt I-J: DefLng L-M: DefLngLng N-O: DefLngPtr P-Q: DefCur C-D: DefSng G-H: DefDbl D-E: DefDate T-U: DefStr E-F: DefObj O-P: DefVar V-W";
            var parseResult = Parse(code, PredictionMode.Sll);
            AssertTree(parseResult.Item1, parseResult.Item2, "//letterRange", matches => matches.Count == 13);
        }

        [Category("Parser")]
        [Test]
        public void TestDefDirectiveUniversalLetterRange()
        {
            string code = @"DefBool A-Z";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//universalLetterRange");
        }

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
        public void TestModuleConfig()
        {
            string code = @"
BEGIN
  MultiUse = -1  'True
END";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleConfigElement");
        }

        [Category("Parser")]
        [Test]
        public void TestVBFormModuleConfig()
        {
            string code = @"
Begin VB.Form Form1 
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
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleConfig", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void TestVBFormWithHexLiteralModuleConfig()
        {
            string code = @"
Begin VB.Form Form1 
   BackColor = &H00FFFFFF&
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
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleConfig", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void TestVBFormWithAbsoluteResourcePathConfig()
        {
            string code = @"
Begin VB.Form Form1 
   BackColor = &H00FFFFFF&
   Caption         =   ""Form1""
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   ""C:\Test\Form1.frx"":0000
   StartUpPosition =   1  'CenterOwner
End
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleConfig", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void TestVBFormWithDnsUncResourcePathConfig()
        {
            string code = @"
Begin VB.Form Form1 
   BackColor = &H00FFFFFF&
   Caption         =   ""Form1""
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   ""\\initech.com\server01\c$\Test\Form1.frx"":0000
   StartUpPosition =   1  'CenterOwner
End
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleConfig", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void TestVBFormWithIPUncResourcePathConfig()
        {
            string code = @"
Begin VB.Form Form1 
   BackColor = &H00FFFFFF&
   Caption         =   ""Form1""
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   ""\\127.0.0.1\Test\Form1.frx"":0000
   StartUpPosition =   1  'CenterOwner
End
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleConfig", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void TestVBFormWithDollarPrependedResourceModuleConfig()
        {
            string code = @"
Begin VB.Form Form1 
   Caption         =   ""Form1""
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   $""Form1.frx"":0000
   StartUpPosition =   1  'CenterOwner
End
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleConfig", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void TestVBFormWithAlphaLeadingHexLiteralResourceOffsetModuleConfig()
        {
            string code = @"
Begin VB.Form Form1 
   Caption         =   ""Form1""
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   ""Form1.frx"":ACBD
   StartUpPosition =   1  'CenterOwner
End
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleConfig", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void TestVBFormWithNumericLeadingHexLiteralResourceOffsetModuleConfig()
        {
            string code = @"
Begin VB.Form Form1 
   Caption         =   ""Form1""
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   ""Form1.frx"":9ABC
   StartUpPosition =   1  'CenterOwner
End
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleConfig", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        [TestCase(@"^A")]
        [TestCase(@"^Z")]
        [TestCase(@"{F1}")]
        [TestCase(@"{F12}")]
        [TestCase(@"^{F1}")]
        [TestCase(@"^{F12}")]
        [TestCase(@"+{F1}")]
        [TestCase(@"+{F12}")]
        [TestCase(@"+^{F1}")]
        [TestCase(@"+^{F12}")]
        [TestCase(@"^{INSERT}")]
        [TestCase(@"+{INSERT}")]
        [TestCase(@"{DEL}")]
        [TestCase(@"+{DEL}")]
        [TestCase(@"%{BKSP}")]
        public void TestVBFormWithMenuShortcutModuleConfig(string shortcut)
        {
            string code = @"
Begin VB.Form Form1 
   Caption         =   ""Form1""
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   ""Form1.frx"":0000
   StartUpPosition =   1  'CenterOwner
   Begin VB.Menu FileMenu 
      Caption         =   ""File""
      Begin VB.Menu FileOpenMenu
         Caption     = ""Open""
         Shortcut    =   " + shortcut + @"
      End
   End 
End
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleConfig", matches => matches.Count == 3);
        }


        [Category("Parser")]
        [Test]
        public void TestNestedVbFormModuleConfig()
        {
            string code = @"
VERSION 5.00
Begin VB.Form Form1
   Caption = ""Main""
   ClientHeight = 2970
   ClientLeft = 60
   ClientTop = 450
   ClientWidth = 8250
   LinkTopic = ""Form1""
   ScaleHeight = 2970
   ScaleWidth = 8250
   StartUpPosition = 2  'CenterScreen
   Begin VB.CommandButton cmdDelete
      Caption = ""Delete""
      Height = 495
      Left = 1320
      TabIndex = 9
      Top = 2280
      Width = 1215
   End
End
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleConfig", matches => matches.Count == 2);
        }

        [Category("Parser")]
        [Test]
        public void TestNestedVbFormModuleConfigWithObjectDeclarations()
        {
            string code = @"
VERSION 5.00
Object = ""{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0""; ""MSADODC.OCX""
Object = ""{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0""; ""MSDATGRD.OCX""
Begin VB.Form Form1
   Caption = ""Main""
   ClientHeight = 2970
   ClientLeft = 60
   ClientTop = 450
   ClientWidth = 8250
   LinkTopic = ""Form1""
   ScaleHeight = 2970
   ScaleWidth = 8250
   StartUpPosition = 2  'CenterScreen
   Begin VB.CommandButton cmdDelete
      Caption = ""Delete""
      Height = 495
      Left = 1320
      TabIndex = 9
      Top = 2280
      Width = 1215
   End
End
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleConfig", matches => matches.Count == 2);
        }

        [Category("Parser")]
        [Test]
        public void TestNestedVbFormModuleConfigWithMultipleChildren()
        {
            string code = @"
VERSION 5.00
Begin VB.Form Form1
   Caption = ""Main""
   ClientHeight = 2970
   ClientLeft = 60
   ClientTop = 450
   ClientWidth = 8250
   LinkTopic = ""Form1""
   ScaleHeight = 2970
   ScaleWidth = 8250
   StartUpPosition = 2  'CenterScreen
   Begin VB.CommandButton cmdDelete
      Caption = ""Delete""
      Height = 495
      Left = 1320
      TabIndex = 9
      Top = 2280
      Width = 1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1
      Bindings = ""frmMain.frx"":0000
      Height = 2055
      Left = 2520
      TabIndex = 0
      Top = 120
      Width = 5655
      _ExtentX = 9975
      _ExtentY = 3625
      _Version = 393216
      HeadLines = 1
      RowHeight = 15
      AllowAddNew = -1  'True
   End
End
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleConfig", matches => matches.Count == 3);
        }

        [Category("Parser")]
        [Test]
        public void TestNestedVbFormModuleConfigWithProperty()
        {
            string code = @"
Begin VB.Form Form1
   Caption = ""Main""
   ClientHeight = 2970
   ClientLeft = 60
   ClientTop = 450
   ClientWidth = 8250
   LinkTopic = ""Form1""
   ScaleHeight = 2970
   ScaleWidth = 8250
   StartUpPosition = 2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1
      Bindings = ""frmMain.frx"":0000
      Height = 2055
      Left = 2520
      TabIndex = 0
      Top = 120
      Width = 5655
      _ExtentX = 9975
      _ExtentY = 3625
      _Version = 393216
      HeadLines = 1
      RowHeight = 15
      AllowAddNew = -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name = ""MS Sans Serif""
         Size = 8.25
         Charset = 0
         Weight = 400
         Underline = 0   'False
         Italic = 0   'False
         Strikethrough = 0   'False
      EndProperty
   End
End
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleConfigProperty", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void TestNestedVbFormModuleConfigWithMultipleProperties()
        {
            string code = @"
VERSION 5.00
Begin VB.Form Form1
   Caption = ""Main""
   ClientHeight = 2970
   ClientLeft = 60
   ClientTop = 450
   ClientWidth = 8250
   LinkTopic = ""Form1""
   ScaleHeight = 2970
   ScaleWidth = 8250
   StartUpPosition = 2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1
      Bindings = ""frmMain.frx"":0000
      Height = 2055
      Left = 2520
      TabIndex = 0
      Top = 120
      Width = 5655
      _ExtentX = 9975
      _ExtentY = 3625
      _Version = 393216
      HeadLines = 1
      RowHeight = 15
      AllowAddNew = -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851}
         Name = ""MS Sans Serif""
         Size = 8.25
         Charset = 0
         Weight = 400
         Underline = 0   'False
         Italic = 0   'False
         Strikethrough = 0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name = ""MS Sans Serif""
         Size = 8.25
         Charset = 0
         Weight = 400
         Underline = 0   'False
         Italic = 0   'False
         Strikethrough = 0   'False
      EndProperty
   End
End
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleConfigProperty", matches => matches.Count == 2);
        }

        [Category("Parser")]
        [Test]
        public void TestNestedVbFormModuleConfigWithNestedProperties()
        {
            string code = @"
VERSION 5.00
Begin VB.Form Form1
   Caption = ""Main""
   ClientHeight = 2970
   ClientLeft = 60
   ClientTop = 450
   ClientWidth = 8250
   LinkTopic = ""Form1""
   ScaleHeight = 2970
   ScaleWidth = 8250
   StartUpPosition = 2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1
      Bindings = ""frmMain.frx"":0000
      Height = 2055
      Left = 2520
      TabIndex = 0
      Top = 120
      Width = 5655
      _ExtentX = 9975
      _ExtentY = 3625
      _Version = 393216
      HeadLines = 1
      RowHeight = 15
      AllowAddNew = -1  'True
      BeginProperty Column00 
         DataField       =   """"
         Caption         =   """"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   """"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
   End
End
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleConfigProperty", matches => matches.Count == 2);
        }

        [Category("Parser")]
        [Test]
        public void TestIndexedProperty()
        {
            string code = @"
VERSION 5.00
Begin VB.Form Form1
   Begin ComctlLib.ListView lvFilter 
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Caption      =   ""ID""
      EndProperty
   End
End
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleConfigProperty", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void TestEmptyComment()
        {
            string code = @"'";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//comment");
        }

        [Category("Parser")]
        [Test]
        public void TestEmptyRemComment()
        {
            string code = @"Rem";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//remComment");
        }

        [Category("Parser")]
        [Test]
        public void TestOneCharRemComment()
        {
            string code = @"Rem a";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//remComment");
        }

        [Category("Parser")]
        [Test]
        public void TestCommentThatLooksLikeAnnotation()
        {
            string code = @"'@param foo; the value of something";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//comment");
        }

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
        public void TestOneCharComment()
        {
            string code = @"'a";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//comment");
        }

        [Category("Parser")]
        [Test]
        public void TestEndEnumMultipleWhiteSpace()
        {
            string code = @"
Enum Test
    anything
End               Enum";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//enumerationStmt");
        }

        [Category("Parser")]
        [Test]
        public void TestEndTypeMultipleWhiteSpace()
        {
            string code = @"
Type Test
    anything As Integer
End             Type";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//udtDeclaration");
        }

        [Category("Parser")]
        [Test]
        public void TestEndFunctionLineContinuation()
        {
            string code = @"
Function Test()

End _
Function";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//functionStmt");
        }

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
        public void TestEndSubroutineLineContinuation()
        {
            string code = @"
Sub Test()

End _
Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//subStmt");
        }

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
        public void TestPropertyGetLineContinuation()
        {
            string code = @"
Property _
Get Test()
End Property";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//propertyGetStmt");
        }

        [Category("Parser")]
        [Test]
        public void TestPropertyLetLineContinuation()
        {
            string code = @"
Property _
Let Test(anything As Integer)
End Property";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//propertyLetStmt");
        }

        [Category("Parser")]
        [Test]
        public void TestPropertySetLineContinuation()
        {
            string code = @"
Property _
Set Test(anything As Application)
End Property";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//propertySetStmt");
        }

        [Category("Parser")]
        [Test]
        public void TestEndPropertyLineContinuation()
        {
            string code = @"
Property Get Test()

End _
Property";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//propertyGetStmt");
        }

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
        public void TestOptionBaseLineContinuation()
        {
            string code = @"
Option _
Base _
1";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleOption");
        }

        [Category("Parser")]
        [Test]
        public void TestOptionExplicitLineContinuation()
        {
            string code = @"
Option _
Explicit";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleOption");
        }

        [Category("Parser")]
        [Test]
        public void TestOptionCompareLineContinuation()
        {
            string code = @"
Option _
Compare _
Text";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleOption");
        }

        [Category("Parser")]
        [Test]
        public void TestOptionPrivateModuleLineContinuation()
        {
            string code = @"
Option _
Private _
Module";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//moduleOption");
        }

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
        public void TestDeclareLineContinuation()
        {
            string code = @"
Private Declare Function ABC Lib ""shell32.dll"" Alias _
""ShellExecuteA""(ByVal a As Long, ByVal b As String, _
ByVal c As String, ByVal d As String, ByVal e As String, ByVal f As Long) As Long";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//declareStmt");
        }

        [Category("Parser")]
        [Test]
        public void TestEraseStmt()
        {
            string code = @"
Public Sub EraseTwoArrays()
Erase someArray(), someOtherArray()
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//eraseStmt");
        }

        [Category("Parser")]
        [Test]
        public void TestFixedLengthString()
        {
            string code = @"
Sub Test()
    Dim someString As String * 255
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//fieldLength");
        }

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
        public void TestCombinedForNextStatement()
        {
            string code = @"
Sub Test()
    For n = 1 To 10
        For m = 1 To 20
            a = m + n
        Next m _
    , n%
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//forNextStmt", matches => matches.Count == 2);
        }

        [Category("Parser")]
        [Test]
        public void TestCombinedForNextStatementWhithItermediateCode()
        {
            string code = @"
Sub Test()
    For n = 1 To 10
        b = n
        For m = 1 To 20
            a = m + n
    Next m,n%
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//forNextStmt", matches => matches.Count == 2);
        }

        [Category("Parser")]
        [Test]
        public void TestCombinedForEachStatement()
        {
            string code = @"
Sub Test()
    Dim foo As Collection
    Dim bar As Collection
    For Each n In foo
        For Each m In bar
            a = m + n
    Next m _
        , _
        n%
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//forEachStmt", matches => matches.Count == 2);
        }

        [Category("Parser")]
        [Test]
        public void TestCombinedForEachStatementWhithItermediateCode()
        {
            string code = @"
Sub Test()
    Dim foo As Collection
    Dim bar As Collection
    For Each n In foo
        b = n
        For Each m In bar
            a = m + n
    Next m,n%
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//forEachStmt", matches => matches.Count == 2);
        }

        [Category("Parser")]
        [Test]
        public void TestMixedCombinedForEachAndForNextStatement()
        {
            string code = @"
Sub Test()
    Dim foo As Collection
    Dim bar As Collection
    For n = 1 To 10
        b = n
        For Each c In foo
            For m = 1 To 20
                For Each d In bar
                    a = m + n + c + d
                        For k = 0 To 100
                            t = a + k
    Next k, d, m, _
            c, _
            n
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//forEachStmt", matches => matches.Count == 2);
            AssertTree(parseResult.Item1, parseResult.Item2, "//forNextStmt", matches => matches.Count == 3);
        }

        [Category("Parser")]
        [Test]
        public void TestMixedRegularAndCombinedForEachAndForNextStatement()
        {
            string code = @"
Sub Test()
    Dim foo As Collection
    Dim bar As Collection
    For n = 1 To 10
        For Each c In foo
        Next c
        For m = 1 To 20
            For k = 0 To 100
                t = a + k
            Next
            For Each d In bar
                For l = 15 To 23
                   a = m + n + d + l
            Next l, d                   
    Next m, n
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//forEachStmt", matches => matches.Count == 2);
            AssertTree(parseResult.Item1, parseResult.Item2, "//forNextStmt", matches => matches.Count == 4);
        }

        [Category("Parser")]
        [Test]
        public void TestLineLabelStatement()
        {
            string code = @"
Sub Test()
a:
10:
154
12 b:
52'comment
644 _

71Rem stupid Rem comment
22 

77 _
 : 
42whatever
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//statementLabelDefinition", matches => matches.Count == 10);
        }

        [Category("Parser")]
        [Test]
        public void TestLineLabelStatementWithCodeOnSameLine()
        {
            string code = @"
Sub Test()
a: foo
10: bar: foo
15 bar
12 b: foo: bar
77 _
 : bar
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//statementLabelDefinition", matches => matches.Count == 5);
            AssertTree(parseResult.Item1, parseResult.Item2, "//callStmt", matches => matches.Count == 7);
        }

        [Category("Parser")]
        [Test]
        public void NameStatement()
        {
            string code = @"
Sub Test()
    Dim sOldPath, sOldName As String
    Dim sNewPath, sNewName As String
    Name sOldPath + sOldName As sNewPath + sNewName
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//nameStmt", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void ProcedureNamedName()
        {
            string code = @"
Sub Name()
End Sub

Sub Test()
    Name
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//identifier", matches => matches.Count == 3);    // name, test, and name
        }

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
        public void TestEmptyAnnotationsWithParentheses()
        {
            string code = @"
'@NoIndent()
Sub Test()
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//annotation");
        }

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
        public void TestSingleLineIfEmptyThenEmptyElse()
        {
            string code = @"
Sub Test()
    If False Then Else
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//singleLineIfStmt");
        }

        [Category("Parser")]
        [Test]
        public void TestSingleLineIfEmptyThenEndOfStatement()
        {
            string code = @"
Sub Test()
    If False Then: Else
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//singleLineIfStmt");
        }

        [Category("Parser")]
        [Test]
        public void TestSingleLineIfMultipleThenNoElse()
        {
            string code = @"
Sub Test()
      If False Then MsgBox False: MsgBox False Else
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//singleLineIfStmt");
        }

        [Category("Parser")]
        [Test]
        public void TestSingleLineIfMultipleThenMultipleElse()
        {
            string code = @"
Sub Test()
      If False Then MsgBox False: MsgBox False Else MsgBox False: MsgBox False
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//singleLineIfStmt");
        }

        [Category("Parser")]
        [Test]
        public void TestSingleLineIfEmptyThen()
        {
            string code = @"
Sub Test()
      If False Then Else MsgBox True
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//singleLineIfStmt");
        }

        [Category("Parser")]
        [Test]
        public void TestSingleLineIfSingleThenEmptyElse()
        {
            string code = @"
Sub Test()
      If False Then MsgBox True Else
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//singleLineIfStmt");
        }

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
        public void TestSingleLineIfDoLoop()
        {
            string code = @"
Sub Test()
      If True Then Do: Loop Else
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//doLoopStmt");
        }

        [Category("Parser")]
        [Test]
        public void TestSingleLineIfWendLoop()
        {
            string code = @"
Sub Test()
      If True Then While True: Beep: Wend Else
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//whileWendStmt");
        }

        [Category("Parser")]
        [Test]
        public void TestSingleLineIfRealWorldExample1()
        {
            string code = @"
Sub Test()
      On Local Error Resume Next: If Not Empty Is Nothing Then Do While Null: ReDim i(True To False) As Currency: Loop: Else Debug.Assert CCur(CLng(CInt(CBool(False Imp True Xor False Eqv True)))): Stop: On Local Error GoTo 0
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//singleLineIfStmt");
        }

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
        public void TestSingleLineIfRealWorldExample3()
        {
            string code = @"
Sub Test()
    If Not oP_Window Is Nothing Then If Not oP_Window.Visible Then Unload oP_Window: Set oP_Window = Nothing
End Sub";
            var parseResult = Parse(code, PredictionMode.Sll);
            AssertTree(parseResult.Item1, parseResult.Item2, "//singleLineIfStmt", matches => matches.Count == 2);
        }

        [Category("Parser")]
        [Test]
        public void TestSingleLineIfRealWorldExample4()
        {
            string code = @"
Sub Test()
    If Err Then Set oP_Window = Nothing: TurnOff Else If oP_Window Is Nothing Then TurnOn
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//singleLineIfStmt", matches => matches.Count == 2);
        }

        [Category("Parser")]
        [Test]
        public void TestEndStmt()
        {
            string code = @"
Sub Test()
    End
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//endStmt");
        }

        [Category("Parser")]
        [Test]
        public void TestRedimStmtArray()
        {
            string code = @"
Sub Test()
    ReDim strArray(1)
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//expression");
        }

        [Category("Parser")]
        [Test]
        public void TestRedimStmtLowerBoundsArgument()
        {
            string code = @"
Sub Test()
    ReDim strArray(1 To 10)
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//lowerBoundArgumentExpression");
        }

        [Category("Parser")]
        [Test]
        public void TestRedimStmtUpperBoundsArgument()
        {
            string code = @"
Sub Test()
    ReDim strArray(1 To 10)
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//upperBoundArgumentExpression");
        }

        [Category("Parser")]
        [Test]
        public void TestRedimStmtNormalArgument()
        {
            string code = @"
Sub Test()
    ReDim strArray(1 To 10)
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//argumentExpression");
        }

        [Category("Parser")]
        [Test]
        public void TestStringFunction()
        {
            string code = @"
Sub Test()
    a = String(5, ""a"")
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//lExpression");
        }

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
        public void TestOpenStmt()
        {
            string code = @"
Sub Test()
    Open ""TESTFILE"" For Binary Access Read Lock Read As #1 Len = 2
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//openStmt");
        }

        [Category("Parser")]
        [Test]
        public void TestResetStmt()
        {
            string code = @"
Sub Test()
    Reset
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//resetStmt");
        }

        [Category("Parser")]
        [Test]
        public void TestCloseStmt()
        {
            string code = @"
Sub Test()
    Close #1, 2, 3
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//closeStmt");
        }

        [Category("Parser")]
        [Test]
        public void TestSeekStmt()
        {
            string code = @"
Sub Test()
    Seek #1, 2
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//seekStmt");
        }

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
        public void TestLockStmt()
        {
            string code = @"
Sub Test()
    Lock #1, 2
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//lockStmt");
        }

        [Category("Parser")]
        [Test]
        public void TestUnlockStmt()
        {
            string code = @"
Sub Test()
    Unlock #1, 2
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//unlockStmt");
        }

        [Category("Parser")]
        [Test]
        public void TestLineInputStmt()
        {
            string code = @"
Sub Test()
    Line Input #2, ""ABC""
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//lineInputStmt");
        }

        [Category("Parser")]
        [Test]
        public void TestWidthStmt()
        {
            string code = @"
Sub Test()
    Width #2, 5
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//widthStmt");
        }

        [Category("Parser")]
        [Test]
        public void TestPrintStmt()
        {
            string code = @"
Sub Test()
    Print #2, Spc(5) ;
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//printStmt");
        }

        [Category("Parser")]
        [Test]
        public void TestDebugPrintStmtNoArguments()
        {
            string code = @"
Sub Test()
    Debug.Print
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//debugPrintStmt");
        }

        [Category("Parser")]
        [Test]
        public void TestDebugPrintStmtNormalArgumentSyntax()
        {
            string code = @"
Sub Test()
    Debug.Print ""Anything""
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//debugPrintStmt/outputList");
        }

        [Category("Parser")]
        [Test]
        public void TestDebugPrintStmtOutputItemSemicolon()
        {
            string code = @"
Sub Test()
    Debug.Print 1;
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//debugPrintStmt/outputList");
        }

        [Category("Parser")]
        [Test]
        public void TestDebugPrintStmtOutputItemComma()
        {
            string code = @"
Sub Test()
    Debug.Print 1,
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//debugPrintStmt/outputList");
        }

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
        public void TestDebugPrintRealWorldExample2()
        {
            string code = @"
Sub Test()
    If Not pFault Then
        Debug.Print ""FirstO: "" & vbCr & ans(0) & vbCr
        Debug.Print ""SecondO:""; ans(1)
    End If
End Sub";
            var parseResult = Parse(code, PredictionMode.Sll);
            AssertTree(parseResult.Item1, parseResult.Item2, "//debugPrintStmt", matches => matches.Count == 2);
        }

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
        public void TestWriteStmt()
        {
            string code = @"
Sub Test()
    Write #1, ""ABC"", 234
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//writeStmt");
        }

        [Category("Parser")]
        [Test]
        public void TestInputStmt()
        {
            string code = @"
Sub Test()
    Input #1, ""ABC""
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//inputStmt");
        }

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
        public void TestCircleSpecialForm()
        {
            string code = @"
Sub Test()
    Me.Circle Step (1, 2), 3, 4, 5, 6, 7
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//circleSpecialForm");
        }

        [Category("Parser")]
        [Test]
        public void TestCircleSpecialForm_WithoutStep()
        {
            string code = @"
Sub Test()
    Me.Circle (1, 2), 3, 4, 5, 6, 7
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//circleSpecialForm");
        }

        [Category("Parser")]
        [Test]
        public void TestCircleSpecialForm_WithoutOptionalArguments()
        {
            string code = @"
Sub Test()
    Me.Circle Step (1, 2), 3
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//circleSpecialForm");
        }

        [Category("Parser")]
        [Test]
        public void TestLineAccessReport()
        {
            string code = @"
Sub Test()
    Me.Line Step(1, 1)-Step(2, 2), vbBlack, B
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//lineSpecialForm");
        }

        [Category("Parser")]
        [Test]
        public void TestLineAccessReport_WithoutOptionalArguments()
        {
            string code = @"
Sub Test()
    Me.Line (1, 1)-(2, 2)
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//lineSpecialForm");
        }

        [Category("Parser")]
        [Test]
        public void TestLineAccessReport_WithoutStartingTuple()
        {
            string code = @"
Sub Test()
    Me.Line -(2, 2)
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//lineSpecialForm");
        }

        [Category("Parser")]
        [Test]
        public void TestLineAccessReport_WithoutStep()
        {
            string code = @"
Sub Test()
    Me.Line (1, 1)-(2, 2), vbBlack, BF
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//lineSpecialForm");
        }

        [Category("Parser")]
        [Test]
        public void TestScaleSpecialForm()
        {
            string code = @"
Sub Test()
    Scale (1, 2)-(3, 4)
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//scaleSpecialForm");
        }

        [Category("Parser")]
        [Test]
        public void TestPSetVBForm_WithoutStep()
        {
            string code = @"
Sub Test()
    Me.PSet (1, 2), vbBlack
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//pSetSpecialForm");
        }

        [Category("Parser")]
        [Test]
        public void TestPSetVBForm_WithoutOptionalArguments()
        {
            string code = @"
Sub Test()
    Me.PSet (1, 2)
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//pSetSpecialForm");
        }

        [Category("Parser")]
        [Test]
        public void TestPSetSpecialForm()
        {
            string code = @"
Sub Test()
    PSet Step(1, 2), vbBlack
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//pSetSpecialForm");
        }

        [Category("Parser")]
        [Test]
        public void TestPSetSpecialForm_WithoutStep()
        {
            string code = @"
Sub Test()
    PSet (1, 2), vbBlack
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//pSetSpecialForm");
        }

        [Category("Parser")]
        [Test]
        public void TestPSetSpecialForm_WithoutOptionalArguments()
        {
            string code = @"
Sub Test()
    PSet (1, 2)
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//pSetSpecialForm");
        }

        [Category("Parser")]
        [Test]
        public void TestPtrSafeAsSub()
        {
            string code = @"
Private Sub PtrSafe()
    Debug.Print 42
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//subStmt");
        }

        [Category("Parser")]
        [Test]
        public void TestFunction_Indented()
        {
            string code = @"
    Private Function Foo() As Boolean
        Foo = True
    End Function";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//functionStmt");
        }

        [Category("Parser")]
        [Test]
        public void TestSub_Indented()
        {
            string code = @"
    Private Sub Foo()
    End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//subStmt");
        }

        [Category("Parser")]
        [Test]
        public void TestSub_InconsistentlyIndented()
        {
            string code = @"
    Private Sub Foo()
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//subStmt");
        }

        [Category("Parser")]
        [Test]
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

        [Category("Parser")]
        [Test]
        public void TestLiteralExpressionResolvesCorrectly()
        {
            string code = @"
Private Sub Foo()
    a = True
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//literalExpression");
        }

        [Category("Parser")]
        [Test]
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
        
        [Category("Parser")]
        [Test]
        public void TestNestedParensForLiteralExpression()
        {
            const string code = @"
Sub Test()
    Dim foo As Integer
    foo = ((42) + ((12)))
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//literalExpression", matches => matches.Count == 2);
        }

        [Category("Parser")]
        [Test]
        public void TestParensForByValSingleArg()
        {
            const string code = @"
Sub Test()
    DoSomething (foo)
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//argumentExpression", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void TestParensForByValFirstArg()
        {
            const string code = @"
Sub Test()
    DoSomething (foo), bar
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//argumentExpression", matches => matches.Count == 2);
        }

        [Category("Parser")]
        [Test]
        public void TestDefaultMemberAccessCallStmtOnFunctionReturnValue_Single()
        {
            const string code = @"
Sub Test()
    SomeFunction(foo) bar
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//callStmt", matches => matches.Count == 1);
            AssertTree(parseResult.Item1, parseResult.Item2, "//argumentExpression", matches => matches.Count == 2);
            AssertTree(parseResult.Item1, parseResult.Item2, "//argumentList", matches => matches.Count == 2);
        }

        [Category("Parser")]
        [Test]
        public void TestDefaultMemberAccessCallStmtOnFunctionReturnValue_Multiple()
        {
            const string code = @"
Sub Test()   
    SomeFunction(foo, bar) foobar
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//callStmt", matches => matches.Count == 1);
            AssertTree(parseResult.Item1, parseResult.Item2, "//argumentExpression", matches => matches.Count == 3);
            AssertTree(parseResult.Item1, parseResult.Item2, "//argumentList", matches => matches.Count == 2);
        }

        [Category("Parser")]
        [Test]
        public void TestFunctionArgumentsOnContinuedLine_Multiple()
        {
            const string code = @"
Sub Test()
    Dim x As Long    
    x = SomeFunction _
    (foo, bar)
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//letStmt", matches => matches.Count == 1);
            AssertTree(parseResult.Item1, parseResult.Item2, "//argumentExpression", matches => matches.Count == 2);
        }

        [Category("Parser")]
        [Test]
        public void TestFunctionArgumentsOnContinuedLine_Single()
        {
            const string code = @"
Sub Test()
    Dim x As Long    
    x = SomeFunction _
    (foo)
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//letStmt", matches => matches.Count == 1);
            AssertTree(parseResult.Item1, parseResult.Item2, "//argumentExpression", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void TestDefaultMemberAccessCallStmtOnFunctionReturnValue_FunctionArgumentsOnContinuedLine_Single()
        {
            const string code = @"
Sub Test() 
    SomeFunction _
    (foo) bar
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//callStmt", matches => matches.Count == 1);
            AssertTree(parseResult.Item1, parseResult.Item2, "//argumentExpression", matches => matches.Count == 2);
            AssertTree(parseResult.Item1, parseResult.Item2, "//argumentList", matches => matches.Count == 2);
        }

        [Category("Parser")]
        [Test]
        public void TestDefaultMemberAccessCallStmtOnFunctionReturnValue_FunctionArgumentsOnContinuedLine_Multiple()
        {
            const string code = @"
Sub Test()   
    SomeFunction _
    (foo, bar) foobar
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//callStmt", matches => matches.Count == 1);
            AssertTree(parseResult.Item1, parseResult.Item2, "//argumentExpression", matches => matches.Count == 3);
            AssertTree(parseResult.Item1, parseResult.Item2, "//argumentList", matches => matches.Count == 2);
        }

        [Category("Parser")]
        [Test]
        public void ParserDoesNotFailOnUnderscoreComment()
        {
            const string code = @"
Sub Test()   
    '_
    If True Then
    End If
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//ifStmt", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void ParserDoesNotFailOnUnderscoreAfterNonBreakingSpaceInComment()
        {
            const string code = @"
Sub Test()   
    '" + "\u00A0" + @"_
    If True Then
    End If
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//ifStmt", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void ParserDoesNotFailOnStartOfLineUnderscoreInLineContinuedComment()
        {
            const string code = @"
Sub Test()   
    ' _
_
    If True Then
    End If
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//ifStmt", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void ParserDoesNotFailOnLineContinuedMemberAccessExpressionInType1()
        {
            const string code = @"
Sub Test()   
Dim dic2 As _
Scripting _
. _
Dictionary
End Sub
";
            var parseResult = Parse(code);
        }

        [Category("Parser")]
        [Test]
        public void ParserDoesNotFailOnLineContinuedMemberAccessExpressionInType2()
        {
            const string code = @"
Sub Test()   
  Dim dic3 As Scripting _
  . _
  Dictionary
End Sub
";
            var parseResult = Parse(code);
        }

        [Category("Parser")]
        [Test]
        public void ParserDoesNotFailOnLineContinuedMemberAccessExpressionOnObject1()
        {
            const string code = @"
Sub Test()   
Dim dict As Scripting.Dictionary

  Debug.Print dict. _
  Item(""a"")
End Sub
";
            var parseResult = Parse(code);
        }

        [Category("Parser")]
        [Test]
        public void ParserDoesNotFailOnLineContinuedMemberAccessExpressionOnObject2()
        {
            const string code = @"
Sub Test()   
Dim dict As Scripting.Dictionary

Debug.Print dict _
. _
Item(""a"")
End Sub
";
            var parseResult = Parse(code);
        }

        [Category("Parser")]
        [Test]
        public void ParserDoesNotFailOnLineContinuedMemberAccessExpressionOnObject3()
        {
            const string code = @"
Sub Test()   
Dim dict As Scripting.Dictionary

Debug.Print dict _
    . _
Item(""a"")
End Sub
";
            var parseResult = Parse(code);
        }

        [Category("Parser")]
        [Test]
        public void ParserDoesNotFailOnLineContinuedBangOperator1()
        {
            const string code = @"
Sub Test()   
Dim dict As Scripting.Dictionary

Dim x
x = dict _
! _
a
End Sub
";
            var parseResult = Parse(code);
        }

        [Category("Parser")]
        [Test]
        public void ParserDoesNotFailOnLineContinuedBangOperator2()
        {
            const string code = @"
Sub Test()   
Dim dict As Scripting.Dictionary

Dim x
x = dict _
  ! _
  a

End Sub
";
            var parseResult = Parse(code);
        }

        [Category("Parser")]
        [Test]
        public void ParserDoesNotFailOnLineContinuedBangOperator3()
        {
            const string code = @"
Sub Test()   
Dim dict As Scripting.Dictionary

Dim x
x = dict _
!a
End Sub
";
            var parseResult = Parse(code);
        }

        [Category("Parser")]
        [Test]
        public void ParserDoesNotFailOnLineContinuedTypeDeclaration()
        {
            const string code = @"
Sub Test()   
Dim dic1 As _
Dictionary
End Sub
";
            var parseResult = Parse(code);
        }

        [Category("Parser")]
        [Test]
        public void ParserDoesNotFailOnIdentifierEndingInUnderscore()
        {
            const string code = @"
Sub Test()   
Dim dict As Scripting.Dictionary

Dim x_
End Sub
";
            var parseResult = Parse(code);
        }

        [Category("Parser")]
        [Test]
        public void ParserDoesNotFailOnLineNumberNotOnStartOfLineAfterALineContinuation()
        {
            const string code = @"
Sub foo()
 _
    10
 _ 
Beep
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//standaloneLineNumberLabel", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void ParserDoesNotFailOnLineLAbelNotOnStartOfLineAfterALineContinuation()
        {
            const string code = @"
Sub foo()
 _
    foo: Beep
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//identifierStatementLabel", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void ParserDoesNotFailOnLinecontinuedLabel()
        {
            const string code = @"
Sub foo()
foo _
: Beep
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//identifierStatementLabel", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void ParserDoesNotFailOnLineNumberAndLineContinuedLabelNotOnStartOfLineAfterALineContinuation()
        {
            const string code = @"
Sub foo()
 _
    10
 _
    foo _
    : Beeb
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//standaloneLineNumberLabel", matches => matches.Count == 1);
            AssertTree(parseResult.Item1, parseResult.Item2, "//identifierStatementLabel", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void ParserDoesNotFailOnLineNumberAndLineContinuedLabelNotOnStartOfLineAfterMultipleLineContinuation()
        {
            const string code = @"
Sub foo()
 _
 _
 _
    10
 _
 _
 _
    foo _
    : Beeb
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//standaloneLineNumberLabel", matches => matches.Count == 1);
            AssertTree(parseResult.Item1, parseResult.Item2, "//identifierStatementLabel", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void LeftOutOptionalArgumentsAreCountedAsMissingArguments()
        {
            const string code = @"
Public Sub Test()
    Dim x As Long
    x = Foo(1, , 5)
End Sub

Public Function Foo(a, Optional b, Optional c) As Long
    Foo = 42
End Function
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//missingArgument", matches => matches.Count == 1);
            AssertTree(parseResult.Item1, parseResult.Item2, "//argumentList", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void TestReDimWithLineContinuation()
        {
            const string code = @"
Sub Test()
    Dim x() As Long    
    Redim x _
    (1, 2)
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//redimStmt", matches => matches.Count == 1);
            AssertTree(parseResult.Item1, parseResult.Item2, "//argumentExpression", matches => matches.Count == 2);
        }

        [Category("Parser")]
        [Test]
        public void TestCaseIsEqExpressionWithLiteral()
        {
            const string code = @"
Sub Test(ByVal foo As Integer)
    Select Case foo
        Case Is = 42
            Exit Sub
    End Select
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//rangeClause", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void TestCaseIsEqExpressionWithEnum()
        {
            const string code = @"
Sub Test(ByVal foo As vbext_ComponentType)
    Select Case foo
        Case Is = vbext_ct_StdModule
            Exit Sub
    End Select
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//rangeClause", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void TestRaiseEventByValParameter()
        {
            const string code = @"
Public Event Foo(ByRef Bar As Boolean, ByVal Baz As String)

Public Sub Test()
    Dim arg As String
    arg = ""Foo""
    RaiseEvent Foo(True, ByVal 42)
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//raiseEventStmt", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void TestRaiseEvent()
        {
            const string code = @"
Public Event Foo(ByRef Bar As Boolean, ByVal Baz As String)

Public Sub Test()
    Dim arg As Boolean
    RaiseEvent Foo(arg, ""Foo"")
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//raiseEventStmt", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void TestAttributeAfterSub()
        {
            const string code = @"
Public Sub Foo(): End Sub
Attribute Foo.VB_Description = ""Foo description""

Public Sub Bar()
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//attributeStmt", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void TestAttributeAfterFunction()
        {
            const string code = @"
Public Function Foo(): End Function
Attribute Foo.VB_Description = ""Foo description""

Public Sub Bar()
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//attributeStmt", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void TestAttributeAfterPropertyGet()
        {
            const string code = @"
Public Property Get Foo(): End Property
Attribute Foo.VB_Description = ""Foo description""

Public Sub Bar()
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//attributeStmt", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void TestAttributeAfterPropertyLet()
        {
            const string code = @"
Public Property Let Foo(): End Property
Attribute Foo.VB_Description = ""Foo description""

Public Sub Bar()
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//attributeStmt", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void TestAttributeAfterPropertySet()
        {
            const string code = @"
Public Property Set Foo(): End Property
Attribute Foo.VB_Description = ""Foo description""

Public Sub Bar()
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//attributeStmt", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void SubtractionExpressionsAreNoLetterRanges()
        {
            const string code = @"
Public Sub Foo()
    Dim a As Long
    Dim b As Long
    Dim z As Long
    a = 1
    b = 2
    z = a-b
    b = a-z
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//letterRange", matches => matches.Count == 0);
            AssertTree(parseResult.Item1, parseResult.Item2, "//universalLetterRange", matches => matches.Count == 0);
        }

        [Category("Parser")]
        [Test]
        public void SLLParserDoesNotThrowForArrayDefinitionInModuleWithMultipleSpacesInFromtOfAsType()
        {
            const string code = @"
Dim Foo1(0 To 3)       As Long

";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//variableSubStmt", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void SLLParserDoesNotThrowForArrayDefinitionInSubWithMultipleSpacesInFromtOfAsType()
        {
            const string code = @"
Sub Test()
    Dim Foo2(0 To 3)       As Long
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//variableSubStmt", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void UserDefinedType_TreatsFinalCommentAsComment()
        {
            // See Issue #3789
            const string code = @"
Private Type tX
    foo As String
    bar As Long
    'foobar as shouldNotBeVisible
End Type
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//udtMember", matches => matches.Count == 2);
            AssertTree(parseResult.Item1, parseResult.Item2, "//commentOrAnnotation", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void MidStatement()
        {
            const string code = @"
Public Sub Test()
    Dim TestString As String
    TestString = ""The dog jumps""
    Mid(TestString, 5, 3) = ""fox""
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//midStatement", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void MidDollarStatement()
        {
            const string code = @"
Public Sub Test()
    Dim TestString As String
    TestString = ""The dog jumps""
    Mid$(TestString, 5, 3) = ""fox""
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//midStatement", matches => matches.Count == 1);
        }

        public void MidBStatement()
        {
            const string code = @"
Public Sub Test()
    Dim TestString As String
    TestString = ""The dog jumps""
    MidB(TestString, 5, 3) = ""fox""
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//midStatement", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void MidBDollarStatement()
        {
            const string code = @"
Public Sub Test()
    Dim TestString As String
    TestString = ""The dog jumps""
    MidB$(TestString, 5, 3) = ""fox""
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//midStatement", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void MidFunction()
        {
            const string code = @"
Public Sub Test()
    Dim TestString As String
    TestString = ""The dog jumps""
    If Mid(TestString, 5, 3) = ""fox"" Then
        MsgBox ""Found""
    End If
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//midStatement", matches => matches.Count == 0);
        }

        [Category("Parser")]
        [Test]
        public void MidDollarFunction()
        {
            const string code = @"
Public Sub Test()
    Dim TestString As String
    TestString = ""The dog jumps""
    If Mid$(TestString, 5, 3) = ""fox"" Then
        MsgBox ""Found""
    End If
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//midStatement", matches => matches.Count == 0);
        }

        [Category("Parser")]
        [Test]
        public void MidBFunction()
        {
            const string code = @"
Public Sub Test()
    Dim TestString As String
    TestString = ""The dog jumps""
    If MidB(TestString, 5, 3) = ""fox"" Then
        MsgBox ""Found""
    End If
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//midStatement", matches => matches.Count == 0);
        }

        [Category("Parser")]
        [Test]
        public void MidBDollarFunction()
        {
            const string code = @"
Public Sub Test()
    Dim TestString As String
    TestString = ""The dog jumps""
    If MidB$(TestString, 5, 3) = ""fox"" Then
        MsgBox ""Found""
    End If
End Sub
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//midStatement", matches => matches.Count == 0);
        }

        [Category("Parser")]
        [Test]
        public void ParserAcceptsScaleMemberInUDT()
        {
            const string code = @"
Public Type Whatever
    Scale As Double
End Type
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//unrestrictedIdentifier", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void ParserAcceptsCircleMemberInUDT()
        {
            const string code = @"
Public Type Whatever
    Circle As Long
End Type
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//unrestrictedIdentifier", matches => matches.Count == 1);
        }

        [Category("Parser")]
        [Test]
        public void ParserAcceptsPSetMemberInUDT()
        {
            const string code = @"
Public Type Whatever
    PSet As Boolean
End Type
";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//unrestrictedIdentifier", matches => matches.Count == 1);
        }

        
        [Test]
        [Category("Parser")]
        [TestCase("Private WithEvents foo As EventSource, WithEvents bar As EventSource", 2)]
        [TestCase("Private WithEvents foo As EventSource, bar As EventSource", 2)]
        [TestCase("Private foo As EventSource, WithEvents bar As EventSource", 2)]
        [TestCase("Private foo As EventSource, bar As EventSource", 2)]
        [TestCase("Private WithEvents foo As EventSource", 1)]
        public void WithEventsInVariableList(string code, int count)
        {
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//variableSubStmt", matches => matches.Count == count);
        }

        private Tuple<VBAParser, ParserRuleContext> Parse(string code, PredictionMode predictionMode = null)
        {
            var stream = new AntlrInputStream(code);
            var lexer = new VBALexer(stream);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBAParser(tokens);
            // Don't remove this line otherwise we won't get notified of parser failures.
            parser.ErrorHandler = new BailErrorStrategy();
            //parser.AddErrorListener(new ExceptionErrorListener());
            parser.Interpreter.PredictionMode = predictionMode ?? PredictionMode.Sll;
            ParserRuleContext tree;
            try
            {
                tree = parser.startRule();
            }
            catch (Exception exception)
            {
                if (predictionMode == null || predictionMode == PredictionMode.Ll)
                {
                    // If SLL fails we want to get notified ASAP so we can fix it, that's why we don't retry using LL.
                    // If LL mode fails, we're done.

                    throw;
                }

                Debug.WriteLine(exception, "SLL Parser Exception");
                return Parse(code, PredictionMode.Ll);
            }
            return Tuple.Create(parser, tree);
        }

        private void AssertTree(VBAParser parser, ParserRuleContext root, string xpath, string message = "")
        {
            AssertTree(parser, root, xpath, matches => matches.Count >= 1, message);
        }

        private void AssertTree(VBAParser parser, ParserRuleContext root, string xpath, Predicate<ICollection<IParseTree>> assertion, string message = "")
        {
            var matches = new XPath(parser, xpath).Evaluate(root);
            var actual = matches.Count;
            Assert.IsTrue(assertion(matches), $"{actual} matches found. {message}");
        }
    }
}