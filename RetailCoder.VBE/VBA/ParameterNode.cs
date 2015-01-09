using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Rubberduck.VBA
{
    [ComVisible(false)]
    public class ParameterNode : SyntaxTreeNode
    {
        public ParameterNode(Instruction instruction, string scope, Match match)
            : base(instruction, scope, match, new[] {new IdentifierNode(instruction, scope, match)})
        {
            _isImplicitByRef = !match.Groups["by"].Success;
        }

        public IdentifierNode Identifier { get { return ChildNodes.OfType<IdentifierNode>().Single(); } }

        private readonly bool _isImplicitByRef;
        public bool IsImplicitByRef { get { return _isImplicitByRef; } }
    }
}