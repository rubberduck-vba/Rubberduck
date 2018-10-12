namespace Rubberduck.Parsing.VBA.Parsing.ParsingExceptions
{
    public class MainParseErrorListenerFactory : IParsePassErrorListenerFactory
    {
        public IRubberduckParseErrorListener Create(string moduleName, CodeKind codeKind)
        {
            return new MainParseExceptionErrorListener(moduleName, codeKind);
        }
    }
}
