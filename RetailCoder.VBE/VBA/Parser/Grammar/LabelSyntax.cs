using System.Text.RegularExpressions;

namespace Rubberduck.VBA.Parser.Grammar
{
    internal class LabelSyntax : SyntaxBase
    {
        protected override bool MatchesSyntax(string instruction, out Match match)
        {
            match = Regex.Match(instruction, VBAGrammar.LabelSyntax);
            return match.Success;
        }

        protected override SyntaxTreeNode CreateNode(Instruction instruction, string scope, Match match)
        {
            return new LabelNode(instruction, scope, match);
        }
    }
}
