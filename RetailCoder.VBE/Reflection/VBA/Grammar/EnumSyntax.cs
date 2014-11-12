using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Rubberduck.Reflection.VBA.Grammar
{
    internal class EnumSyntax : ISyntax
    {
        public SyntaxTreeNode ToNode(string publicScope, string localScope, Instruction instruction)
        {
            var match = Regex.Match(instruction.Content, VBAGrammar.EnumSyntax());
            if (!match.Success)
            {
                return null;
            }

            return new EnumNode(instruction, localScope, match);
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

    internal class EnumMemberSyntax : ISyntax
    {
        public SyntaxTreeNode ToNode(string publicScope, string localScope, Instruction instruction)
        {
            var match = Regex.Match(instruction.Content.Trim(), VBAGrammar.EnumMemberSyntax());
            if (!match.Success)
            {
                return null;
            }

            return new EnumMemberNode(instruction, localScope, match);
        }

        public bool IsMatch(string publicScope, string localScope, Instruction instruction, out SyntaxTreeNode node)
        {
            node = ToNode(publicScope, localScope, instruction);
            return node != null;
        }

        public bool IsChildNodeSyntax
        {
            get { return true; }
        }
    }
}
