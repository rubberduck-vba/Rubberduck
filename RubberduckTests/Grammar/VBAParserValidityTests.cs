using NUnit.Framework;
using Rubberduck.Parsing.VBA;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor;

namespace RubberduckTests.Grammar
{
    [TestFixture]
    public class VBAParserValidityTests
    {
        [Test]
        [Category("LongRunning")]
        [Category("Grammar")]
        [Category("Parser")]
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
            var basePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            return Directory.EnumerateFiles(Path.Combine(basePath, "Testfiles//Grammar")).Select(file => Tuple.Create(file, File.ReadAllText(file))).ToList();
        }

        private static string Parse(string code, string filename)
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out component);

            string parsedCode;
            var parser = MockParser.Create(vbe.Object);
            using (var state = parser.State)
            {
                parser.Parse(new CancellationTokenSource());

                if (state.Status == ParserState.Error)
                {
                    Assert.Inconclusive("Parser Error: " + filename);
                }

                var tree = state.GetParseTree(new QualifiedModuleName(component));
                parsedCode = tree.GetText();
            }
            var withoutEOF = parsedCode;
            while (withoutEOF.Length >= 5 && String.Equals(withoutEOF.Substring(withoutEOF.Length - 5, 5), "<EOF>"))
            {
                withoutEOF = withoutEOF.Substring(0, withoutEOF.Length - 5);
            }
            return withoutEOF;
        }
    }
}
