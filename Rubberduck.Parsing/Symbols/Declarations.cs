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

        public IEnumerable<Declaration> this[DeclarationType declarationType, string identifierName]
        {
            get
            {
                return _declarations.Where(declaration =>
                            declaration.DeclarationType == declarationType &&
                            declaration.IdentifierName == identifierName);
            }
        }
    }
}
