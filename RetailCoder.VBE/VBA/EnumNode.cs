using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA
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

        public string Accessibility { get { return GetAccessibility(); } }

        private string GetAccessibility()
        {
            var keywords = new[] { ReservedKeywords.Public, ReservedKeywords.Private, ReservedKeywords.Friend };
            var value = RegexMatch.Groups["accessibility"].Value;

            return (keywords.Contains(value)) ? value : ReservedKeywords.Public;
        }
    }
}
