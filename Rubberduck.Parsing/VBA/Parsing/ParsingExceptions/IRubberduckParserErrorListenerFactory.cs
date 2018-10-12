namespace Rubberduck.Parsing.VBA.Parsing.ParsingExceptions
{
    public interface IRubberduckParserErrorListenerFactory
    {
        IRubberduckParseErrorListener Create(CodeKind codeKind);
    }
}
