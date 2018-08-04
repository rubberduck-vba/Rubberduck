using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;

namespace Rubberduck.Parsing.Symbols.ParsingExceptions
{
    public interface IParsePassErrorListenerFactory
    {
        IRubberduckParseErrorListener Create(string moduleName, CodeKind codeKind);
    }
}
