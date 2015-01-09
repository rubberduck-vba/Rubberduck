using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Rubberduck.VBA.Grammar
{
    [ComVisible(false)]
    public class IfBlockSyntax : SyntaxBase
    {
        public IfBlockSyntax()
            : base(SyntaxType.HasChildNodes)
        {
            
        }

        protected override bool MatchesSyntax(string instruction, out Match match)
        {
            match = Regex.Match(instruction, VBAGrammar.IfBlockSyntax);
            return match.Success;
        }

        protected override SyntaxTreeNode CreateNode(Instruction instruction, string scope, Match match)
        {
            return new IfBlockNode(instruction, scope, match, new List<SyntaxTreeNode>());
        }
    }
}
