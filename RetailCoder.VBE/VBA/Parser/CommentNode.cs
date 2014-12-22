using System.Runtime.InteropServices;

namespace Rubberduck.VBA.Parser
{
    [ComVisible(false)]
    public class CommentNode : SyntaxTreeNode
    {
        public CommentNode(Instruction instruction, string scope)
            : base(instruction, scope, null, null)
        {
        }

        public string Comment { get { return Instruction.Comment; } }
    }
}