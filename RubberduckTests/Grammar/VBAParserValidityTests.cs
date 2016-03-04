using Antlr4.Runtime;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace RubberduckTests.Grammar
{
    [TestClass]
    public class VBAParserValidityTests
    {
        [TestMethod]
        [DeploymentItem(@"Testfiles\")]
        public void TestTest()
        {
            foreach (var testfile in GetTestFiles())
            {
                AssertParseResult(testfile, Parse(testfile).module().GetText());
            }
        }

        private void AssertParseResult(string originalCode, string materializedParseTree)
        {
            Assert.AreEqual(originalCode, materializedParseTree);
        }

        private IEnumerable<string> GetTestFiles()
        {
            return Directory.EnumerateFiles("Grammar").Select(file => File.ReadAllText(file)).ToList();
        }

        private VBAParser.StartRuleContext Parse(string code)
        {
            var stream = new AntlrInputStream(code);
            var lexer = new VBALexer(stream);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBAParser(tokens);
            parser.AddErrorListener(new ExceptionErrorListener());
            return parser.startRule();
        }
    }
}
