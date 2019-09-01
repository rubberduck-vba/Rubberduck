using System.Collections.Generic;
using Antlr4.Runtime.Tree;

namespace Rubberduck.Inspections.CodePathAnalysis.Nodes
{
    public class AssignmentNode : NodeBase
    {
        public AssignmentNode(IParseTree tree) : base(tree) { }

        private readonly IList<INode> _usages = new List<INode>();
        public IEnumerable<INode> Usages => _usages;

        public void AddUsage(INode node) => _usages.Add(node);
    }
}
