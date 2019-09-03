using System.Collections.Generic;
using Antlr4.Runtime.Tree;

namespace Rubberduck.Inspections.CodePathAnalysis.Nodes
{
    public class AssignmentNode : NodeBase
    {
        public AssignmentNode(IParseTree tree) 
            : base(tree)
        {
        }

        private readonly HashSet<INode> _usages = new HashSet<INode>();

        /// <summary>
        /// Gets all nodes reading this assignment's value.
        /// </summary>
        public IEnumerable<INode> Usages => _usages;

        internal void AddUsage(INode node) => _usages.Add(node);
    }
}
