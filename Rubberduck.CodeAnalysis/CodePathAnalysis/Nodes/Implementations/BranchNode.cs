using Antlr4.Runtime.Tree;

namespace Rubberduck.Inspections.CodePathAnalysis.Nodes
{
    public class BranchNode : NodeBase, IBranchNode
    {
        public BranchNode(IParseTree tree) : base(tree) { }
    }
}
