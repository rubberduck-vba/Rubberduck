using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Annotations;

namespace Rubberduck.Parsing.Grammar
{
    public interface IIdentifierContext
    {
        Interval IdentifierTokens { get; }
    }
}
