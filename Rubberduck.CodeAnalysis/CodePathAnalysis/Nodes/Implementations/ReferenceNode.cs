using Antlr4.Runtime.Tree;

namespace Rubberduck.Inspections.CodePathAnalysis.Nodes
{
    public class ReferenceNode : NodeBase
    {
        public ReferenceNode(IParseTree tree, bool isConditional) : base(tree)
        {
            IsConditional = isConditional;
        }

        public bool IsConditional { get; }
    }
}
