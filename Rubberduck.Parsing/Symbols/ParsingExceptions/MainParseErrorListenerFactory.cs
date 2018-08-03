using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing.Symbols.ParsingExceptions
{
    public class MainParseErrorListenerFactory : IParsePassErrorListenerFactory
    {
        public IRubberduckParseErrorListener Create(string moduleName, CodeKind codeKind)
        {
            return new MainParseExceptionErrorListener(moduleName, codeKind);
        }
    }
}
