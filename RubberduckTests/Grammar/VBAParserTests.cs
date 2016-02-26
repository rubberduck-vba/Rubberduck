using Antlr4.Runtime;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace RubberduckTests.Grammar
{
    [TestClass]
    public class VBAParserTests
    {
        [TestMethod]
        public void TestTrivialCase()
        {
            string code = @":";
            string expectedTree = @"
(startRule
    (module
        (endOfStatement :)
        endOfStatement
        endOfStatement
        endOfStatement
        endOfStatement)
<EOF>)";
            assertTree(code, expectedTree);
        }

        [TestMethod]
        public void TestModuleHeader()
        {
            string code = @"VERSION 1.0 CLASS";
            string expectedTree = @"
(startRule
    (module endOfStatement
        (moduleHeader VERSION 1.0 CLASS)
        endOfStatement
        endOfStatement
        endOfStatement
        endOfStatement
        endOfStatement)
<EOF>)";
            assertTree(code, expectedTree);
        }

        [TestMethod]
        public void TestModuleConfig()
        {
            string code = @"
BEGIN
  MultiUse = -1  'True
END";
            string expectedTree = @"
(startRule
    (module
        (endOfStatement
            (endOfLine \r\n))
        (moduleConfig BEGIN
            (endOfStatement
                (endOfLine \r\n ))
            (moduleConfigElement
                (ambiguousIdentifier MultiUse) = 
                (literal -1)
                (endOfStatement
                    (endOfLine
                        (comment 'True))
                    (endOfLine \r\n)))
        END)
        endOfStatement
        endOfStatement
        endOfStatement
        endOfStatement)
<EOF>)";
            assertTree(code, expectedTree);
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
            string expectedTree = @"
(startRule
    (module
        (endOfStatement
            (endOfLine \r\n))
        endOfStatement
        endOfStatement
        endOfStatement
        (moduleBody
            (moduleBodyElement
                (subStmt
                    (visibility Public)
                    Sub
                    (ambiguousIdentifier Test)
                    (argList ( ))
                    (endOfStatement
                        (endOfLine \r\n ))
                    (block
                        (blockStmt
                            (setStmt
                            Set
                            (implicitCallStmt_InStmt
                                (iCS_S_VariableOrProcedureCall
                                    (ambiguousIdentifier result))) =
                            (valueStmt 
                                (implicitCallStmt_InStmt
                                    (iCS_S_VariableOrProcedureCall
                                        (ambiguousIdentifier myObj) _\r\n
                                        (dictionaryCallStmt ! _\r\n
                                            (ambiguousIdentifier
                                            (ambiguousKeyword call))))))))
                    (endOfStatement (endOfLine \r\n)))
                    End Sub))
                endOfStatement)
            endOfStatement)
<EOF>)";
            assertTree(code, expectedTree);
        }

        private void assertTree(string code, string expectedTree)
        {
            var stream = new AntlrInputStream(code);
            var lexer = new VBALexer(stream);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBAParser(tokens);
            parser.AddErrorListener(new ExceptionErrorListener());
            var actualTree = parser.startRule().ToStringTree(parser);
            actualTree = Regex.Replace(actualTree, @"\s+", " ");
            var lines = expectedTree
                .Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries)
                .Select(line => line.Trim());
            var clean = string.Join(" ", lines);
            Assert.AreEqual(clean, actualTree);
        }
    }
}
