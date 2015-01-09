using System.Runtime.InteropServices;

namespace Rubberduck.VBA
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