using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;

namespace Rubberduck.Parsing.Symbols.ParsingExceptions
{
    public interface IRubberduckParserErrorListenerFactory
    {
        IRubberduckParseErrorListener Create(CodeKind codeKind);
    }
}
