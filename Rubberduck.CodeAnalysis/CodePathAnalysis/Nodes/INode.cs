using Rubberduck.Parsing.Symbols;
using System.Collections.Immutable;

namespace Rubberduck.Inspections.CodePathAnalysis.Nodes
{
    public interface INode
    {
        int SortOrder { get; set; }
        ImmutableList<INode> Children { get; set; }
        INode Parent { get; set; }

        Declaration Declaration { get; set; }
        IdentifierReference Reference { get; set; }
    }
}
