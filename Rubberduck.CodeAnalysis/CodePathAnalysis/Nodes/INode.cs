using Rubberduck.Parsing.Symbols;
using System.Collections.Immutable;
using Antlr4.Runtime.Tree;

namespace Rubberduck.Inspections.CodePathAnalysis.Nodes
{
    public interface INode
    {
        int SortOrder { get; set; }
        ImmutableList<INode> Children { get; set; }
        INode Parent { get; set; }
        IParseTree ParseTree { get; }
        Declaration Declaration { get; set; }
        IdentifierReference Reference { get; set; }
    }
}
