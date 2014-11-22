using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Rubberduck.VBA.Parser.Grammar
{
    [ComVisible(false)]
    public class EnumMemberSyntax : SyntaxBase
    {
        public EnumMemberSyntax()
            : base(SyntaxType.IsChildNodeSyntax)
        { 
        }

        protected override bool MatchesSyntax(string instruction, out Match match)
        {
            match = Regex.Match(instruction, VBAGrammar.EnumMemberSyntax);
            return match.Success;
        }

        protected override SyntaxTreeNode CreateNode(Instruction instruction, string scope, Match match)
        {
            return new EnumMemberNode(instruction, scope, match);
        }
    }
}