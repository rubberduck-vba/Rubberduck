namespace Rubberduck.Parsing.VBA.Parsing.ParsingExceptions
{
    public interface IParsePassErrorListenerFactory
    {
        IRubberduckParseErrorListener Create(string moduleName, CodeKind codeKind);
    }
}
