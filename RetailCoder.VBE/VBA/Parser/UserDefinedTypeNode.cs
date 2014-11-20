using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Rubberduck.VBA.Parser.Grammar;

namespace Rubberduck.VBA.Parser
{
    [ComVisible(false)]
    public class UserDefinedTypeNode : CodeBlockNode
    {
        public UserDefinedTypeNode(Instruction instruction, string scope, Match match) 
            : base(instruction, scope, match, new []{string.Concat(ReservedKeywords.End, " ", ReservedKeywords.Type)}, typeof(UserDefinedTypeMemberSyntax), new List<UserDefinedTypeMemberNode>())
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