using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Rubberduck.Reflection.VBA.Grammar
{
    internal class ProcedureSyntax : SyntaxBase
    {
        public ProcedureSyntax()
            : base(SyntaxType.HasChildNodes)
        {

        }

        protected override bool MatchesSyntax(string instruction, out Match match)
        {
            match = Regex.Match(instruction, VBAGrammar.ProcedureSyntax());
            return match.Success;
        }

        protected override SyntaxTreeNode CreateNode(string scope, Match match, string instruction, string comment)
        {
            return new ProcedureNode(scope, match, comment);
        }
    }
}
