﻿using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.VBEHost;
using RubberduckTests.Mocks;
using System;
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
        public void TestParser()
        {
            foreach (var testfile in GetTestFiles())
            {
                var filename = testfile.Item1;
                var code = testfile.Item2;
                AssertParseResult(filename, code, Parse(code));
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

        private string Parse(string code)
        {
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(code, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var state = new RubberduckParserState();
            var parser = new RubberduckParser(vbe.Object, state);
            parser.ParseSynchronous();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }
            var tree = state.GetParseTree(component);
            var parsed = tree.GetText();
            var withoutEOF = parsed.Substring(0, parsed.Length - 5);
            return withoutEOF;
        }
    }
}
