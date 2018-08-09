using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;

namespace Rubberduck.Parsing.Symbols.ParsingExceptions
{
    public class ExceptionErrorListenerFactory : IRubberduckParserErrorListenerFactory
    {
        public IRubberduckParseErrorListener Create(CodeKind codeKind)
        {
            return new ExceptionErrorListener(codeKind);
        }
    }
}
