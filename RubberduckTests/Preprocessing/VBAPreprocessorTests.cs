using Microsoft.VisualStudio.TestTools.UnitTesting;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor;

namespace RubberduckTests.PreProcessing
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
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out component);
            
            var state = MockParser.CreateAndParse(vbe.Object);
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
