using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RetailCoderVBE.Reflection.VBA.Grammar
{
    internal interface ISyntax
    {
        SyntaxTreeNode ToNode(string instruction);
    }

    internal class DimSyntax : ISyntax
    {
        public SyntaxTreeNode ToNode(string instruction)
        {
            var pattern = VBAGrammar.GetLocalDeclarationSyntax(ReservedKeywords.Dim);

            var match = Regex.Match(instruction, pattern);
            if (!match.Success)
            {
                return null;
            }

            return new VariableNode(match);
        }
    }

    internal class PrivateFieldSyntax : ISyntax
    {
        public SyntaxTreeNode ToNode(string instruction)
        {
            var pattern = VBAGrammar.GetModuleDeclarationSyntax(ReservedKeywords.Private);

            var match = Regex.Match(instruction, pattern);
            if (!match.Success)
            {
                return null;
            }

            return new VariableNode(match);
        }
    }

    internal class PublicFieldSyntax : ISyntax
    {
        public SyntaxTreeNode ToNode(string instruction)
        {
            var pattern = VBAGrammar.GetModuleDeclarationSyntax(ReservedKeywords.Public);

            var match = Regex.Match(instruction, pattern);
            if (!match.Success)
            {
                return null;
            }

            return new VariableNode(match);
        }
    }

    internal class GlobalFieldSyntax : ISyntax
    {
        public SyntaxTreeNode ToNode(string instruction)
        {
            var pattern = VBAGrammar.GetModuleDeclarationSyntax(ReservedKeywords.Global);

            var match = Regex.Match(instruction, pattern);
            if (!match.Success)
            {
                return null;
            }

            return new VariableNode(match);
        }
    }

    internal class LocalConstantSyntax : ISyntax
    {
        public SyntaxTreeNode ToNode(string instruction)
        {
            var pattern = VBAGrammar.GetLocalConstantDeclarationSyntax();

            var match = Regex.Match(instruction, pattern);
            if (!match.Success)
            {
                return null;
            }

            return new ConstantNode(match);
        }
    }
}
