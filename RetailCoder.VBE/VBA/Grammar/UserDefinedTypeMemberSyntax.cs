using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Rubberduck.VBA.Grammar
{
    [ComVisible(false)]
    public class UserDefinedTypeMemberSyntax : SyntaxBase
    {
        public UserDefinedTypeMemberSyntax()
            : base(SyntaxType.IsChildNodeSyntax)
        { }

        protected override bool MatchesSyntax(string instruction, out Match match)
        {
            match = Regex.Match(instruction, VBAGrammar.IdentifierDeclarationSyntax);
            return match.Success;
        }

        protected override SyntaxTreeNode CreateNode(Instruction instruction, string scope, Match match)
        {
            return new UserDefinedTypeMemberNode(instruction, scope, match);
        }
    }
}