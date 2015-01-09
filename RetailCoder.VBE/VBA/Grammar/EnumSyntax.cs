using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Rubberduck.VBA.Grammar
{
    [ComVisible(false)]
    public class EnumSyntax : SyntaxBase
    {
        public EnumSyntax()
            : base(SyntaxType.HasChildNodes)
        { }

        protected override bool MatchesSyntax(string instruction, out Match match)
        {
            match = Regex.Match(instruction, VBAGrammar.EnumSyntax);
            return match.Success;
        }

        protected override SyntaxTreeNode CreateNode(Instruction instruction, string scope, Match match)
        {
            return new EnumNode(instruction, scope, match);
        }
    }
}
