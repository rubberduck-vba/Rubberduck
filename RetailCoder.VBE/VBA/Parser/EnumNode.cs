using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Rubberduck.VBA.Parser.Grammar;

namespace Rubberduck.VBA.Parser
{
    [ComVisible(false)]
    public class EnumNode : CodeBlockNode
    {
        public EnumNode(Instruction instruction, string scope, Match match)
            : base(instruction, scope, match, new[] {string.Concat(ReservedKeywords.End, " ", ReservedKeywords.Enum)}, typeof(EnumMemberSyntax), new List<EnumMemberNode>())
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
