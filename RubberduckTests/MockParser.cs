using System;
using System.Collections.Generic;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using RubberduckTests.Mocks;
using Rubberduck.Parsing.Preprocessing;
using System.Globalization;
using System.Threading;
using Rubberduck.Parsing;

namespace RubberduckTests
{
    public static class MockParser
    {
        public static void ParseString(string inputCode, out QualifiedModuleName qualifiedModuleName, out RubberduckParserState state)
        {

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            qualifiedModuleName = new QualifiedModuleName(component);
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object, new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }
            state = parser.State;

        }
        public static RubberduckParser Create(VBE vbe, RubberduckParserState state)
        {
            var attributeParser = new Mock<IAttributeParser>();
            attributeParser.Setup(m => m.Parse(It.IsAny<VBComponent>()))
                           .Returns(() => new Dictionary<Tuple<string, DeclarationType>, Attributes>());
            return Create(vbe, state, attributeParser.Object);
        }

        public static RubberduckParser Create(VBE vbe, RubberduckParserState state, IAttributeParser attributeParser)
        {
            return new RubberduckParser(state, attributeParser,
                () => new VBAPreprocessor(double.Parse(vbe.Version, CultureInfo.InvariantCulture)),
                new List<ICustomDeclarationLoader> {new DebugDeclarations(state), new FormEventDeclarations(state)});
        }
    }
}
