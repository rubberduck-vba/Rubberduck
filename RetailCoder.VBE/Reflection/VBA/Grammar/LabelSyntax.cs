using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Rubberduck.Reflection.VBA.Grammar
{
    internal class LabelSyntax : ISyntax
    {
        public SyntaxTreeNode ToNode(string publicScope, string localScope, Instruction instruction)
        {
            var match = Regex.Match(instruction.Value, VBAGrammar.LabelSyntax());
            if (match.Success)
            {
                return new LabelNode(instruction, localScope, match);
            }

            return null;
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
