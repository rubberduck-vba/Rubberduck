﻿using Antlr4.Runtime;
using Antlr4.Runtime.Tree.Xpath;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;

namespace RubberduckTests.Grammar
{
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
