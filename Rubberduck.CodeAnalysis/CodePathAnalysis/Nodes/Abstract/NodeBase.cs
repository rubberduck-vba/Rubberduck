using System.Collections.Generic;
using System.Collections.Immutable;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.CodePathAnalysis.Nodes
{
    public abstract class NodeBase : INode
    {
        protected NodeBase(IParseTree tree)
        {
            Children = new List<INode>().ToImmutableList();
            ParseTree = tree;
        }

        public int SortOrder { get; set; }
        public ImmutableList<INode> Children { get; set; }
        public INode Parent { get; set; }
        public IParseTree ParseTree { get; }
        public Declaration Declaration { get; set; }
        public IdentifierReference Reference { get; set; }

        public override bool Equals(object obj) =>
            obj is INode node && node.ParseTree.Equals(ParseTree);

        public override int GetHashCode() => HashCode.Compute(ParseTree);
    }
}
