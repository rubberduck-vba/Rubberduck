using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.VBA;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

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
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(code, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var state = new RubberduckParserState(vbe.Object);
            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error: " + filename); }
            var tree = state.GetParseTree(component);
            var parsed = tree.GetText();
            var withoutEOF = parsed.Substring(0, parsed.Length - 5);
            return withoutEOF;
        }
    }
}
