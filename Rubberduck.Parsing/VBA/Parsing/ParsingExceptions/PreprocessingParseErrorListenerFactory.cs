namespace Rubberduck.Parsing.VBA.Parsing.ParsingExceptions
{
    public class PreprocessingParseErrorListenerFactory : IParsePassErrorListenerFactory
    {
        public IRubberduckParseErrorListener Create(string moduleName, CodeKind codeKind)
        {
            return new PreprocessorExceptionErrorListener(moduleName, codeKind);
        }
    }
}
