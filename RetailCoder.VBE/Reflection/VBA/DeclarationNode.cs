using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Rubberduck.Reflection.VBA.Grammar;

namespace Rubberduck.Reflection.VBA
{
    /// <summary>
    /// Base class for a declaration node.
    /// </summary>
    [ComVisible(false)]
    public abstract class DeclarationNode : SyntaxTreeNode
    {
        protected DeclarationNode(Instruction instruction, string scope, Match match)
            : base(instruction, scope, match, true)
        {
            _identifierNodes = ParseIdentifierNodes(instruction, scope, match);
        }

        private readonly IEnumerable<IdentifierNode> _identifierNodes;
        public IEnumerable<IdentifierNode> IdentifierNodes { get { return _identifierNodes; } }

        /// <summary>
        /// Gets the declared identifiers.
        /// </summary>
        /// <example>
        /// Returns all identifiers in a declaration instruction.
        /// </example>
        private IEnumerable<IdentifierNode> ParseIdentifierNodes(Instruction instruction, string scope, Match match)
        {
            var identifiers = match.Groups["expression"].Value.Split(',')
                .Select(identifier => identifier.Trim());

            foreach (var identifier in identifiers)
            {
                var declaration = match.Groups["keywords"].Value + ' ' + identifier;
                var subMatch = Regex.Match(declaration, VBAGrammar.DeclarationKeywordsSyntax() + VBAGrammar.IdentifierDeclarationSyntax());
                yield return new IdentifierNode(instruction, scope, subMatch);
            }
        }
    }

    [ComVisible(false)]
    public class ConstDeclarationNode : DeclarationNode
    {
        public ConstDeclarationNode(Instruction instruction, string scope, Match match) 
            : base(instruction, scope, match)
        {
        }
    }

    [ComVisible(false)]
    public class VariableDeclarationNode : DeclarationNode
    {
        public VariableDeclarationNode(Instruction instruction, string scope, Match match) 
            : base(instruction, scope, match)
        {
        }
    }
}
