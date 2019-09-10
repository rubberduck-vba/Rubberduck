using Antlr4.Runtime.Misc;

namespace Rubberduck.Parsing.Grammar.Abstract
{
    public interface IIdentifierContext
    {
        Interval IdentifierTokens { get; }
    }
}
