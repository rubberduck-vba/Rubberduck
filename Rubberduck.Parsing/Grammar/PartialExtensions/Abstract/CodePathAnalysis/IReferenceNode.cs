using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis
{
    /// <summary>
    /// A node that represents a reference to a declaration.
    /// </summary>
    public interface IReferenceNode : IExtendedNode
    {
        IdentifierReference Reference { get; set; }
    }
}