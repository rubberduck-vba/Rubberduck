using System.Collections.Generic;
using System.Collections.Immutable;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.CodePathAnalysis.Nodes
{
    public class GenericNode : INode
    {
        public GenericNode()
        {
            Children = new List<INode>().ToImmutableList();
        }

        public int SortOrder { get; set; }
        public ImmutableList<INode> Children { get; set; }
        public INode Parent { get; set; }

        public Declaration Declaration { get; set; }
        public IdentifierReference Reference { get; set; }
    }
}
