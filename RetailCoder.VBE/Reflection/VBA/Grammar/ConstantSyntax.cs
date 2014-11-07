using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RetailCoderVBE.Reflection.VBA.Grammar
{
    internal class ConstantSyntax : ISyntax
    {
        public SyntaxTreeNode ToNode(string instruction)
        {
            var pattern = VBAGrammar.GetConstantDeclarationSyntax();

            var match = Regex.Match(instruction, pattern);
            if (!match.Success)
            {
                return null;
            }

            return new ConstantNode(match);
        }
    }
}
