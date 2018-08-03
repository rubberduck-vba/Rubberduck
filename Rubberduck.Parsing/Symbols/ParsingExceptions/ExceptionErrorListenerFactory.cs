using Rubberduck.Parsing.VBA;

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
