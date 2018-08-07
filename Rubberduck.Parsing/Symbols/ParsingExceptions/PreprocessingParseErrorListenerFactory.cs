using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;

namespace Rubberduck.Parsing.Symbols.ParsingExceptions
{
    public class PreprocessingParseErrorListenerFactory : IParsePassErrorListenerFactory
    {
        public IRubberduckParseErrorListener Create(string moduleName, CodeKind codeKind)
        {
            return new PreprocessorExceptionErrorListener(moduleName, codeKind);
        }
    }
}
