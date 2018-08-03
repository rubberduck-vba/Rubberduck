using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing.Symbols.ParsingExceptions
{
    public interface IRubberduckParserErrorListenerFactory
    {
        IRubberduckParseErrorListener Create(CodeKind codeKind);
    }
}
