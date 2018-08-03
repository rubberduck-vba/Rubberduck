using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing.Symbols.ParsingExceptions
{
    public interface IParsePassErrorListenerFactory
    {
        IRubberduckParseErrorListener Create(string moduleName, CodeKind codeKind);
    }
}
