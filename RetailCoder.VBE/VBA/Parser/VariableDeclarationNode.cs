using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Rubberduck.VBA.Parser.Grammar;

namespace Rubberduck.VBA.Parser
{
    [ComVisible(false)]
    public class VariableDeclarationNode : DeclarationNode
    {
        public VariableDeclarationNode(Instruction instruction, string scope, Match match, IEnumerable<IdentifierNode> childNodes) 
            : base(instruction, scope, match, childNodes)
        {
        }

        /// <summary>
        /// Gets the declared identifiers.
        /// </summary>
        /// <example>
        /// Returns all identifiers in a declaration instruction.
        /// </example>
        public static IEnumerable<IdentifierNode> ParseChildNodes(Instruction instruction, string scope, Match match)
        {
            var identifiers = Regex.Match(match.Groups["expression"].Value, VBAGrammar.IdentifierDeclarationSyntax());

            foreach (var identifier in identifiers.Groups["declarations"].Captures)
            {
                var value = identifier.ToString();
                if (value.Trim().EndsWith(","))
                {
                    value = value.Substring(0, value.LastIndexOf(','));
                }

                var declaration = match.Groups["keywords"].Value + ' ' + value;
                var pattern = VBAGrammar.DeclarationKeywordsSyntax() + VBAGrammar.IdentifierDeclarationSyntax();
                var subMatch = Regex.Match(declaration, pattern);
                yield return new IdentifierNode(instruction, scope, subMatch);
            }
        }
    }
}