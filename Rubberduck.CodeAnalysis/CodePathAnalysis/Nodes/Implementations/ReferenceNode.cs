using Antlr4.Runtime.Tree;

namespace Rubberduck.Inspections.CodePathAnalysis.Nodes
{
    public class ReferenceNode : NodeBase
    {
        public ReferenceNode(IParseTree tree, AssignmentNode value = null) : base(tree)
        {
            ValueNode = value;
        }

        public AssignmentNode ValueNode { get; set; }
    }
}
