using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RetailCoderVBE.Reflection.VBA.Grammar
{
    internal class LabelSyntax : ISyntax
    {
        public SyntaxTreeNode ToNode(string publicScope, string localScope, string instruction)
        {
            var comment = string.Empty;
            int index;
            if (instruction.HasComment(out index))
            {
                comment = instruction.Substring(index);
            }

            var match = Regex.Match(instruction, VBAGrammar.LabelSyntax());
            if (match.Success)
            {
                return new LabelNode(localScope, match, comment);
            }

            return null;
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
