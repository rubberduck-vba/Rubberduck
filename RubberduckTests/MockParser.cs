using Microsoft.Vbe.Interop;
using Moq;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests
{
    public static class MockParser
    {
        public static RubberduckParser Create(VBE vbe, RubberduckParserState state)
        {
            var attributeParser = new Mock<IAttributeParser>();
            return Create(vbe, state, attributeParser.Object);
        }

        public static RubberduckParser Create(VBE vbe, RubberduckParserState state, IAttributeParser attributeParser)
        {
            return new RubberduckParser(vbe, state, attributeParser);
        }
    }
}