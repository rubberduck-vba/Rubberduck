using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Rubberduck.Reflection.VBA.Grammar
{
    internal class ConstantSyntax : SyntaxBase
    {
        public ConstantSyntax()
            : base(SyntaxType.HasChildNodes)
        {

        }

        protected override bool MatchesSyntax(string instruction, out Match match)
        {
            match = Regex.Match(instruction, VBAGrammar.ConstantDeclarationSyntax());
            return match.Success;
        }

        protected override SyntaxTreeNode CreateNode(string scope, Match match, string instruction, string comment)
        {
            return new ConstantNode(scope, match, comment);
        }
    }
}
