using Antlr4.Runtime.Misc;

namespace Rubberduck.Parsing.Grammar
{
    public interface IIdentifierContext
    {
        Interval IdentifierTokens { get; }
    }
}
