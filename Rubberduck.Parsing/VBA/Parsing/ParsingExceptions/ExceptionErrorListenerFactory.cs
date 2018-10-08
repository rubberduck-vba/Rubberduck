namespace Rubberduck.Parsing.VBA.Parsing.ParsingExceptions
{
    public class ExceptionErrorListenerFactory : IRubberduckParserErrorListenerFactory
    {
        public IRubberduckParseErrorListener Create(CodeKind codeKind)
        {
            return new ExceptionErrorListener(codeKind);
        }
    }
}
