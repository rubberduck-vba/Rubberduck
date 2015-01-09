using System.Runtime.InteropServices;

namespace Rubberduck.VBA.Grammar
{
    [ComVisible(false)]
    public class ExpressionNode : SyntaxTreeNode
    {
        public ExpressionNode(Instruction instruction, string scope)
            : base(instruction, scope)
        {
            
        }
    }
}