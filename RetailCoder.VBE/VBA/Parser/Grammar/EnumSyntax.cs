using System.Text.RegularExpressions;

namespace Rubberduck.VBA.Parser.Grammar
{
    internal class EnumSyntax : SyntaxBase
    {
        public EnumSyntax()
            : base(SyntaxType.HasChildNodes)
        { }

        protected override bool MatchesSyntax(string instruction, out Match match)
        {
            match = Regex.Match(instruction, VBAGrammar.EnumSyntax());
            return match.Success;
        }

        protected override SyntaxTreeNode CreateNode(Instruction instruction, string scope, Match match)
        {
            return new EnumNode(instruction, scope, match);
        }
    }
}
