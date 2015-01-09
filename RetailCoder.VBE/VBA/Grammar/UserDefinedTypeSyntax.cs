using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Rubberduck.VBA.Grammar
{
    [ComVisible(false)]
    public class UserDefinedTypeSyntax : SyntaxBase
    {
        public UserDefinedTypeSyntax()
            : base(SyntaxType.HasChildNodes)
        { }

        protected override bool MatchesSyntax(string instruction, out Match match)
        {
            match = Regex.Match(instruction, VBAGrammar.UserDefinedTypeSyntax);
            return match.Success;
        }

        protected override SyntaxTreeNode CreateNode(Instruction instruction, string scope, Match match)
        {
            return new UserDefinedTypeNode(instruction, scope, match);
        }
    }
}