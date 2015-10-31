using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    public class Declarations : IEnumerable<Declaration>
    {
        private readonly ConcurrentBag<Declaration> _declarations = new ConcurrentBag<Declaration>();


        public static readonly DeclarationType[] PropertyTypes =
        {
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };

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


        public IEnumerator<Declaration> GetEnumerator()
        {
            return _declarations.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _declarations.GetEnumerator();
        }
    }
}
