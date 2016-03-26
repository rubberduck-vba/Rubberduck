using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.VBEHost;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using RubberduckTests.Inspections;

namespace RubberduckTests.Preprocessing
{
    [TestClass]
    public class VBAPreprocessorTests
    {
        [TestMethod]
        [DeploymentItem(@"Testfiles\")]
        public void TestPreprocessor()
        {
            foreach (var testfile in GetTestFiles())
            {
                var filename = testfile.Item1;
                var code = testfile.Item2;
                var expectedProcessed = testfile.Item3;
                var actualProcessed = Parse(code);
                AssertParseResult(filename, expectedProcessed, actualProcessed);
            }
        }

        private void AssertParseResult(string filename, string originalCode, string materializedParseTree)
        {
            Assert.AreEqual(originalCode, materializedParseTree, string.Format("{0} mismatch detected.", filename));
        }

        private IEnumerable<Tuple<string, string, string>> GetTestFiles()
        {
            // Reference_Module_1 = raw, unprocessed code.
            // Reference_Module_1_Processed = result of preprocessor.
            var all = Directory.EnumerateFiles("Preprocessor").ToList();
            var rawAndProcessed = all
                .Where(file => !file.Contains("_Processed"))
                .Select(file => Tuple.Create(file, all.First(f => f.Contains(Path.GetFileNameWithoutExtension(file)) && f.Contains("_Processed")))).ToList();
            return rawAndProcessed
                .Select(file =>
                    Tuple.Create(file.Item1, File.ReadAllText(file.Item1), File.ReadAllText(file.Item2))).ToList();
        }

        private string Parse(string code)
        {
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(code, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var state = new RubberduckParserState();
            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status == ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }
            var tree = state.GetParseTree(component);
            var parsed = tree.GetText();
            var withoutEOF = parsed.Substring(0, parsed.Length - 5);
            return withoutEOF;
        }
    }
}
