using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Rubberduck.VBA.Grammar
{
    [ComVisible(false)]
    public class OptionSyntax : SyntaxBase
    {
        protected override bool MatchesSyntax(string instruction, out Match match)
        {
            match = Regex.Match(instruction, VBAGrammar.OptionSyntax);
            return match.Success;
        }

        protected override SyntaxTreeNode CreateNode(Instruction instruction, string scope, Match match)
        {
            return new OptionNode(instruction, scope, match);
        }
    }
}