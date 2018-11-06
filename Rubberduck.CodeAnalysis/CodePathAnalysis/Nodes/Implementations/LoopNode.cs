using Antlr4.Runtime.Tree;

namespace Rubberduck.Inspections.CodePathAnalysis.Nodes
{
    public class LoopNode : NodeBase, ILoopNode
    {
        public LoopNode(IParseTree tree) : base(tree) { }
    }
}
