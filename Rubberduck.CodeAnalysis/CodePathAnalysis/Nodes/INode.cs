using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using Antlr4.Runtime.Tree;

namespace Rubberduck.Inspections.CodePathAnalysis.Nodes
{
    public interface INode
    {
        int SortOrder { get; set; }
        IReadOnlyList<INode> Children { get; set; }
        IEnumerable<TNode> Ancestors<TNode>();
        IEnumerable<TNode> Descendants<TNode>();
        INode Parent { get; set; }
        IParseTree ParseTree { get; }
        Declaration Declaration { get; set; }
        IdentifierReference Reference { get; set; }
    }

    public interface IExecutableNode : INode
    {
        bool HasExecuted { get; }

        /// <summary>
        /// Simulates execution of the node.
        /// </summary>
        void Execute(ExecutionContext context);
    }
}
