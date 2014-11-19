using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Rubberduck.VBA.Parser.Grammar
{
    [ComVisible(false)]
    public class DeclarationSyntax : SyntaxBase
    {
        protected override bool MatchesSyntax(string instruction, out Match match)
        {
            var reserved = new[]
            {
                ReservedKeywords.Sub,
                ReservedKeywords.Function,
                ReservedKeywords.Property,
                ReservedKeywords.Enum,
                ReservedKeywords.Type,
                ReservedKeywords.Declare
            };

            match = Regex.Match(instruction, VBAGrammar.GeneralDeclarationSyntax);
            var m = match; // out parameter cannot be used in anonymous method body

            return m.Success 
                && m.Groups["keywords"].Success
                && !reserved.Any(keyword => m.Groups["expression"].Value.Contains(keyword));
        }

        private static readonly IDictionary<string, Func<Instruction, string, Match, SyntaxTreeNode>> Factory = 
            new Dictionary<string, Func<Instruction, string, Match, SyntaxTreeNode>>
            {
                { ReservedKeywords.Const, (instruction, scope, match) => new ConstDeclarationNode(instruction, scope, match, ConstDeclarationNode.ParseChildNodes(instruction, scope, match)) },
                { ReservedKeywords.Dim, (instruction, scope, match) => new VariableDeclarationNode(instruction, scope, match, VariableDeclarationNode.ParseChildNodes(instruction, scope, match)) },
                { ReservedKeywords.Public, (instruction, scope, match) => new VariableDeclarationNode(instruction, scope, match, VariableDeclarationNode.ParseChildNodes(instruction, scope, match)) },
                { ReservedKeywords.Private, (instruction, scope, match) => new VariableDeclarationNode(instruction, scope, match, VariableDeclarationNode.ParseChildNodes(instruction, scope, match)) },
                { ReservedKeywords.Global, (instruction, scope, match) => new VariableDeclarationNode(instruction, scope, match, VariableDeclarationNode.ParseChildNodes(instruction, scope, match)) },
                { ReservedKeywords.Friend, (instruction, scope, match) => new VariableDeclarationNode(instruction, scope, match, VariableDeclarationNode.ParseChildNodes(instruction, scope, match)) }
            };

        protected override SyntaxTreeNode CreateNode(Instruction instruction, string scope, Match match)
        {
            var keyword = match.Groups["keywords"].Value.Split(' ').Last();
            var declaration = Factory.ContainsKey(keyword)
                                ? Factory[keyword]
                                : null;

            var result = declaration == null
                            ? null
                            : declaration(instruction, scope, match);

            return result;
        }
    }
}
