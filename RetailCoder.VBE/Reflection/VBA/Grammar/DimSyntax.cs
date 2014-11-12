using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Rubberduck.Reflection.VBA.Grammar
{
    internal class DimSyntax : ISyntax
    {
        public SyntaxTreeNode ToNode(string publicScope, string localScope, Instruction instruction)
        {
            var pattern = VBAGrammar.LocalDeclarationSyntax(ReservedKeywords.Dim);
            var reserved = new[] { ReservedKeywords.Sub, ReservedKeywords.Property, ReservedKeywords.Function, ReservedKeywords.Enum, ReservedKeywords.Type };

            var match = Regex.Match(instruction.Content, pattern);
            if (!match.Success || reserved.Any(keyword => keyword == match.Groups["identifier"].Captures[0].Value))
            {
                return null;
            }

            return new VariableNode(instruction, localScope, match);
        }


        public bool IsMatch(string publicScope, string localScope, Instruction instruction, out SyntaxTreeNode node)
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
