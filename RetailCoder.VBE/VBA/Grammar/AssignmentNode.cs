using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Rubberduck.VBA.Grammar
{
    [ComVisible(false)]
    public class AssignmentNode : SyntaxTreeNode
    {
        public AssignmentNode(Instruction instruction, string scope, Match match = null) 
            : base(instruction, scope, match, null)
        {
        }

        public Identifier Identifier
        {
            get { return new Identifier(Scope, RegexMatch.Groups["identifier"].Value, string.Empty); }
        }

        public Expression Expression
        {
            get {  return new Expression(RegexMatch.Groups["expression"].Value);}
        }
    }
}