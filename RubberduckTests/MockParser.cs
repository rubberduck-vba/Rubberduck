using System;
using System.Collections.Generic;
using Microsoft.Vbe.Interop;
using Moq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests
{
    public static class MockParser
    {
        public static RubberduckParser Create(VBE vbe, RubberduckParserState state)
        {
            var attributeParser = new Mock<IAttributeParser>();
            attributeParser.Setup(m => m.Parse(It.IsAny<VBComponent>()))
                           .Returns(() => new Dictionary<Tuple<string, DeclarationType>, Attributes>());
            return Create(vbe, state, attributeParser.Object);
        }

        public static RubberduckParser Create(VBE vbe, RubberduckParserState state, IAttributeParser attributeParser)
        {
            return new RubberduckParser(vbe, state, attributeParser);
        }
    }
}
