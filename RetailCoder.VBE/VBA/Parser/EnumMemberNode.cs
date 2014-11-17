using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Rubberduck.VBA.Parser
{
    [ComVisible(false)]
    public class EnumMemberNode : SyntaxTreeNode
    {
        public EnumMemberNode(Instruction instruction, string scope, Match match)
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

        public string Value
        {
            get
            {
                var value = RegexMatch.Groups["value"];
                return value.Success 
                    ? value.Value 
                    : string.Empty;
            }
        }
    }
}