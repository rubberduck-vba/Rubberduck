using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.CodePathAnalysis.Nodes
{
    public abstract class NodeBase : INode
    {
        protected NodeBase(IParseTree tree)
        {
            Children = new List<INode>().ToList();
            ParseTree = tree;
        }

        public int SortOrder { get; set; }
        public IReadOnlyList<INode> Children { get; set; }
        public INode Parent { get; set; }
        public IParseTree ParseTree { get; }
        public Declaration Declaration { get; set; }
        public IdentifierReference Reference { get; set; }

        public IEnumerable<TNode> Ancestors<TNode>()
        {
            if (Parent is TNode node)
            {
                yield return node;
                foreach (var ancestor in Parent?.Ancestors<TNode>() ?? Enumerable.Empty<TNode>())
                {
                    yield return ancestor;
                }
            }
        }

        public IEnumerable<TNode> Descendants<TNode>()
        {
            foreach (var child in Children ?? Enumerable.Empty<INode>())
            {
                if (child is TNode node)
                {
                    yield return node;
                    foreach (var descendant in child.Descendants<TNode>())
                    {
                        yield return descendant;
                    }
                }
            }
        }

        public override bool Equals(object obj) =>
            obj is INode node && node.ParseTree.Equals(ParseTree);

        public override int GetHashCode() => HashCode.Compute(ParseTree);
    }
}
