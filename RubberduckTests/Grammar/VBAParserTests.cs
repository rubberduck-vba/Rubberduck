using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Antlr4.Runtime.Tree.Xpath;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace RubberduckTests.Grammar
{
    [TestClass]
    public class AttributeParserTests
    {
        [TestMethod]
        public void ParsesEmptyForm()
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
            var stream = new AntlrInputStream(code);
            var lexer = new VBALexer(stream);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBAParser(tokens);
            parser.ErrorListeners.Clear();
            parser.ErrorListeners.Add(new ExceptionErrorListener());
            var tree = parser.startRule();
            Assert.IsNotNull(tree);
        }
    }

    [TestClass]
    public class VBAParserTests
    {
        [TestMethod]
        public void TestTrivialCase()
        {
            string code = @":";
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
        public void TestOneCharComment()
        {
            string code = @"'a";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//comment");
        }

        [TestMethod]
        public void TestDictionaryCallLineContinuation()
        {
            string code = @"
Public Sub Test()
    Set result = myObj _
    ! _
    call
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//dictionaryCallStmt");
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
        public void TestMemberCallLineContinuation()
        {
            string code = @"
Public Sub Test()
    Debug.Print Foo.Bar _
                   . _
                    FooBar.Baz
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//iCS_S_MembersCall");
        }

        [TestMethod]
        public void TestMemberProcedureCallLineContinuation()
        {
            string code = @"
Sub Test()
	fun(1) _
	.fun(2) _
	.fun(3)
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//iCS_B_MemberProcedureCall");
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
        public void TestStringFunction()
        {
            string code = @"
Sub Test()
    a = String(5, ""a"")
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//iCS_S_VariableOrProcedureCall", matches => matches.Count == 2);
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
            AssertTree(parseResult.Item1, parseResult.Item2, "//implicitCallStmt_InStmt");
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
        public void TestDebugPrintStmt()
        {
            // Sanity check so that we don't break Debug.Print because of the Print statement.
            string code = @"
Sub Test()
    Debug.Print ""Anything""
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//implicitCallStmt_InBlock");
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
            AssertTree(parseResult.Item1, parseResult.Item2, "//implicitCallStmt_InStmt");
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
            AssertTree(parseResult.Item1, parseResult.Item2, "//implicitCallStmt_InStmt");
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

        private Tuple<VBAParser, ParserRuleContext> Parse(string code)
        {
            var stream = new AntlrInputStream(code);
            var lexer = new VBALexer(stream);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBAParser(tokens);
            var root = parser.startRule();
            var k = root.ToStringTree(parser);
            return Tuple.Create<VBAParser, ParserRuleContext>(parser, root);
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
