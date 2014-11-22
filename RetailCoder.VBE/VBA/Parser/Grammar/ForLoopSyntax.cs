using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Rubberduck.VBA.Parser.Grammar
{
    [ComVisible(false)]
    public class ForLoopSyntax : SyntaxBase
    {
        public ForLoopSyntax()
            : base(SyntaxType.HasChildNodes)
        {
            
        }

        protected override bool MatchesSyntax(string instruction, out Match match)
        {
            match = Regex.Match(instruction, VBAGrammar.ForLoopSyntax);
            return match.Success;
        }

        protected override SyntaxTreeNode CreateNode(Instruction instruction, string scope, Match match)
        {
            
        }
    }
}