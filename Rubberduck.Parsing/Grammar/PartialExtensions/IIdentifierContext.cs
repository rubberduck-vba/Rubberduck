using Antlr4.Runtime.Misc;

// ReSharper disable once CheckNamespace
namespace Rubberduck.Parsing.Grammar
{
    public interface IIdentifierContext
    {
        Interval IdentifierTokens { get; }
    }
}
