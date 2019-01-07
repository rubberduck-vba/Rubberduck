using Antlr4.Runtime.Tree;

namespace Rubberduck.Inspections.CodePathAnalysis.Nodes
{
    public class AssignmentNode : NodeBase
    {
        public AssignmentNode(IParseTree tree) : base(tree) { }
    }
}
