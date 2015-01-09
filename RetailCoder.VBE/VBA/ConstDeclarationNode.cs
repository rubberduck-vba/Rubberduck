using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA
{
    [ComVisible(false)]
    public class ConstDeclarationNode : DeclarationNode
    {
        public ConstDeclarationNode(Instruction instruction, string scope, Match match, IEnumerable<IdentifierNode> chilNodes) 
            : base(instruction, scope, match, chilNodes)
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
            var identifiers = match.Groups["expression"].Value.Split(',')
                .Select(identifier => identifier.Trim());

            foreach (var identifier in identifiers)
            {
                var declaration = match.Groups["keywords"].Value + ' ' + identifier;
                var subMatch = Regex.Match(declaration, VBAGrammar.DeclarationKeywordsSyntax + VBAGrammar.IdentifierDeclarationSyntax);
                yield return new IdentifierNode(instruction, scope, subMatch);
            }
        }
    }
}