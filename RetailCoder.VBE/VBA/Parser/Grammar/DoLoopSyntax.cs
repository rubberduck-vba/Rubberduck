using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Rubberduck.VBA.Parser.Grammar
{
    [ComVisible(false)]
    public class DoLoopSyntax : SyntaxBase
    {
        protected override bool MatchesSyntax(string instruction, out Match match)
        {
            match = Regex.Match(instruction, VBAGrammar.DoLoopSyntax);
            return match.Success;
        }

        protected override SyntaxTreeNode CreateNode(Instruction instruction, string scope, Match match)
        {
            return new DoLoopNode(instruction, scope, match, new List<SyntaxTreeNode>());
        }
    }
}