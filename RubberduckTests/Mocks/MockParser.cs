using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Preprocessing;
using System.Globalization;
using System.Threading;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Mocks
{
    public static class MockParser
    {
        public static void ParseString(string inputCode, out QualifiedModuleName qualifiedModuleName, out RubberduckParserState state)
        {

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            qualifiedModuleName = new QualifiedModuleName(component);
            var parser = Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }
            state = parser.State;

        }

        public static ParseCoordinator Create(IVBE vbe, RubberduckParserState state)
        {
            var attributeParser = new Mock<IAttributeParser>();
            attributeParser.Setup(m => m.Parse(It.IsAny<IVBComponent>()))
                           .Returns(() => new Dictionary<Tuple<string, DeclarationType>, Attributes>());
            return Create(vbe, state, attributeParser.Object);
        }

        public static ParseCoordinator Create(IVBE vbe, RubberduckParserState state, IAttributeParser attributeParser)
        {
            return new ParseCoordinator(vbe, state, attributeParser,
                () => new VBAPreprocessor(double.Parse(vbe.Version, CultureInfo.InvariantCulture)),
                new List<ICustomDeclarationLoader>
                {
                    new DebugDeclarations(state), 
                    new SpecialFormDeclarations(state), 
                    new FormEventDeclarations(state), 
                    new AliasDeclarations(state),
                }, true);
        }
    }
}
