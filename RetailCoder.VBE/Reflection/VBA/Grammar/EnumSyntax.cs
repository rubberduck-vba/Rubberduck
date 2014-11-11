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
        public SyntaxTreeNode ToNode(string publicScope, string localScope, string instruction)
        {
            var match = Regex.Match(instruction, VBAGrammar.EnumSyntax());
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

            return new EnumNode(localScope, match, comment);
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

    internal class EnumMemberSyntax : ISyntax
    {
        public SyntaxTreeNode ToNode(string publicScope, string localScope, string instruction)
        {
            var match = Regex.Match(instruction.Trim(), VBAGrammar.EnumMemberSyntax());
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

            return new EnumMemberNode(localScope, match, comment);
        }

        public bool IsMatch(string publicScope, string localScope, string instruction, out SyntaxTreeNode node)
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
