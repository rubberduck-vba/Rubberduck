using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Symbols
{
    public static class DeclarationEnumerableExtensions
    {
        /// <summary>
        /// Gets all declarations of the specified <see cref="DeclarationType"/>.
        /// </summary>
        public static IEnumerable<Declaration> OfType(this IEnumerable<Declaration> declarations, DeclarationType declarationType)
        {
            return declarations.Where(declaration => declaration.DeclarationType.HasFlag(declarationType));
        }

        /// <summary>
        /// Gets the declaration for all identifiers declared in or below the specified scope.
        /// </summary>
        public static IEnumerable<Declaration> InScope(this IEnumerable<Declaration> declarations, Declaration parent)
        {
            return declarations.Where(declaration => declaration.ParentScope == parent.Scope);
        }
    }
}