using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RetailCoderVBE.Reflection.VBA.Grammar
{
    internal class ConstantSyntax : ISyntax
    {
        public SyntaxTreeNode ToNode(string publicScope, string localScope, string instruction)
        {
            var pattern = VBAGrammar.GetConstantDeclarationSyntax();

            var beforeComment = instruction.StripTrailingComment();
            if (string.IsNullOrEmpty(beforeComment))
            {
                return null;
            }

            var match = Regex.Match(beforeComment, pattern);
            if (!match.Success)
            {
                return null;
            }

            var comment = string.Empty;
            int commentStart;
            if (instruction.HasComment(out commentStart))
            {
                comment = instruction.Substring(commentStart);
            }

            var scope = new[] { ReservedKeywords.Public, ReservedKeywords.Global }.Contains(match.Groups[0].Value) ? publicScope
                                                                                                                   : localScope;
            return new ConstantNode(scope, match, comment);
        }
    }
}
