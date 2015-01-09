using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Rubberduck.VBA.Grammar
{
    [ComVisible(false)]
    public class ProcedureSyntax : SyntaxBase
    {
        public ProcedureSyntax()
            : base(SyntaxType.HasChildNodes)
        {
        }

        protected override bool MatchesSyntax(string instruction, out Match match)
        {
            match = Regex.Match(instruction, VBAGrammar.ProcedureSyntax);
            return match.Success;
        }

        protected override SyntaxTreeNode CreateNode(Instruction instruction, string scope, Match match)
        {
            return new ProcedureNode(instruction, scope, match, match.Groups["kind"].Value.Split(' ')[0], new List<SyntaxTreeNode>());
        }
    }
}
