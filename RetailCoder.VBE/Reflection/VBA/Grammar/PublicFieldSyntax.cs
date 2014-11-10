using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RetailCoderVBE.Reflection.VBA.Grammar
{
    internal class PublicFieldSyntax : ISyntax
    {
        public SyntaxTreeNode ToNode(string publicScope, string localScope, string instruction)
        {
            var pattern = VBAGrammar.ModuleDeclarationSyntax(ReservedKeywords.Public);
            var reserved = new[] { ReservedKeywords.Sub, ReservedKeywords.Property, ReservedKeywords.Function, ReservedKeywords.Enum, ReservedKeywords.Type };

            var match = Regex.Match(instruction, pattern);
            if (!match.Success || reserved.Any(keyword => keyword == match.Groups["identifier"].Captures[0].Value))
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

        public bool IsMatch(string publicScope, string localScope, string instruction, out SyntaxTreeNode node)
        {
            node = ToNode(publicScope, localScope, instruction);
            return node != null;
        }


        public bool IsChildNodeSyntax
        {
            get { return false; }
        }
    }
}
