using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RetailCoderVBE.Reflection.VBA.Grammar
{
    internal class GlobalFieldSyntax : ISyntax
    {
        public SyntaxTreeNode ToNode(string publicScope, string localScope, string instruction)
        {
            var pattern = VBAGrammar.GetModuleDeclarationSyntax(ReservedKeywords.Global);

            var match = Regex.Match(instruction, pattern);
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

            return new VariableNode(publicScope, match, comment);
        }
    }
}
