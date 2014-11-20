using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Rubberduck.VBA.Parser
{
    [ComVisible(false)]
    public class UserDefinedTypeMemberNode : SyntaxTreeNode
    {
        public UserDefinedTypeMemberNode(Instruction instruction, string scope, Match match) 
            : base(instruction, scope, match)
        {
        }

        public Identifier Identifier
        {
            get
            {
                var name = RegexMatch.Groups["identifier"].Value;
                return new Identifier(Scope, name, name);
            }
        }
    }
}