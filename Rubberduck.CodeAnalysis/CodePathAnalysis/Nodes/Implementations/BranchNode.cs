using Antlr4.Runtime.Tree;

namespace Rubberduck.Inspections.CodePathAnalysis.Nodes
{
    public class BranchNode : StatementNode, IBranchNode
    {
        public BranchNode(IParseTree tree) : base(tree) { }
    }
}
