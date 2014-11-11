using Rubberduck.Reflection.VBA.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Rubberduck.Reflection.VBA
{
    internal class EnumNode : CodeBlockNode
    {
        public EnumNode(string scope, Match match)
            : this(scope, match, string.Empty)
        {

        }

        public EnumNode(string scope, Match match, string comment)
            : base(scope, match, comment, string.Concat(ReservedKeywords.End, " ", ReservedKeywords.Enum), typeof(EnumMemberSyntax))
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

    internal class EnumMemberNode : SyntaxTreeNode
    {
        public EnumMemberNode(string scope, Match match, string comment)
            : base(scope, match, comment)
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
                if (value.Success)
                {
                    return value.Value;
                }

                return string.Empty;
            }
        }
    }
}
