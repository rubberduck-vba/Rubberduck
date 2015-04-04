using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Symbols
{
    public class Declarations
    {
        private readonly ConcurrentBag<Declaration> _declarations = new ConcurrentBag<Declaration>();

        /// <summary>
        /// Adds specified declaration to available lookups.
        /// </summary>
        /// <param name="declaration">The declaration to add.</param>
        public void Add(Declaration declaration)
        {
            _declarations.Add(declaration);
        }

        public IEnumerable<Declaration> Items { get { return _declarations; } }

        public IEnumerable<Declaration> this[string identifierName]
        {
            get
            {
                return _declarations.Where(declaration =>
                    declaration.IdentifierName == identifierName);
            }
        }

        /// <summary>
        /// Finds all members declared under the scope defined by the specified declaration.
        /// </summary>
        public IEnumerable<Declaration> FindMembers(Declaration parent)
        {
            return _declarations.Where(declaration => declaration.ParentScope == parent.Scope);
        }
    }
}
