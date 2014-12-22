using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Rubberduck.VBA.Parser.Grammar
{
    [ComVisible(true)]
    public class AssignmentSyntax : SyntaxBase
    {
        protected override bool MatchesSyntax(string instruction, out Match match)
        {
            match = Regex.Match(instruction, VBAGrammar.AssignmentSyntax);
            return !Regex.IsMatch(instruction, VBAGrammar.GeneralDeclarationSyntax) && match.Success;
        }

        protected override SyntaxTreeNode CreateNode(Instruction instruction, string scope, Match match)
        {
            return new AssignmentNode(instruction, scope, match);
        }
    }
}