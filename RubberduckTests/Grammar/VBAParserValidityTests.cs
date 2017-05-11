using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.VBA;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor;

namespace RubberduckTests.Grammar
{
    [TestClass]
    public class VBAParserValidityTests
    {
        [TestMethod]
        [TestCategory("LongRunning")]
        [TestCategory("Grammar")]
        [DeploymentItem(@"Testfiles\")]
        public void TestParser()
        {
            foreach (var testfile in GetTestFiles())
            {
                var filename = testfile.Item1;
                var code = testfile.Item2;
                AssertParseResult(filename, code, Parse(code, filename));
            }
        }

        private void AssertParseResult(string filename, string originalCode, string materializedParseTree)
        {
            Assert.AreEqual(originalCode, materializedParseTree, string.Format("{0} mismatch detected.", filename));
        }

        private IEnumerable<Tuple<string, string>> GetTestFiles()
        {
            return Directory.EnumerateFiles("Grammar").Select(file => Tuple.Create(file, File.ReadAllText(file))).ToList();
        }

        private static string Parse(string code, string filename)
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out component);

            var state = new RubberduckParserState(vbe.Object);
            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error: " + filename); }

            var tree = state.GetParseTree(new QualifiedModuleName(component));
            var parsed = tree.GetText();
            var withoutEOF = parsed;
            while (withoutEOF.Length >= 5 && String.Equals(withoutEOF.Substring(withoutEOF.Length - 5, 5), "<EOF>"))
            {
                withoutEOF = withoutEOF.Substring(0, withoutEOF.Length - 5);
            }
            return withoutEOF;
        }
    }
}
