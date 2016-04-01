using Antlr4.Runtime;
using Antlr4.Runtime.Tree.Xpath;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;

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
        public void TestDeleteSettingsStatement()
        {
            string code = @"
Sub Test()
    DELETESETTING ""a""
    DELETESETTING ""a"", ""b""
    DELETESETTING ""a"", ""b"", ""c""
End Sub";
            var parseResult = Parse(code);
            AssertTree(parseResult.Item1, parseResult.Item2, "//deleteSettingStmt");
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
            AssertTree(parseResult.Item1, parseResult.Item2, "//lineLabel");
        }

        private Tuple<VBAParser, ParserRuleContext> Parse(string code)
        {
            var stream = new AntlrInputStream(code);
            var lexer = new VBALexer(stream);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBAParser(tokens);
            parser.AddErrorListener(new ExceptionErrorListener());
            var root = parser.startRule();
            // Useful for figuring out what XPath to use for querying the parse tree.
            var str = root.ToStringTree(parser);
            return Tuple.Create<VBAParser, ParserRuleContext>(parser, root);
        }

        private void AssertTree(VBAParser parser, ParserRuleContext root, string xpath)
        {
            var matches = new XPath(parser, xpath).Evaluate(root);
            Assert.IsTrue(matches.Count >= 1);
        }
    }
}
