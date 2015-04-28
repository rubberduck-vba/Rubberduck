using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;

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

        /// <summary>
        /// Finds all event handler procedures for specified control declaration.
        /// </summary>
        public IEnumerable<Declaration> FindEventHandlers(Declaration control)
        {
            return _declarations.Where(declaration => declaration.ParentScope == control.ParentScope
                && declaration.DeclarationType == DeclarationType.Procedure
                && declaration.IdentifierName.StartsWith(control.IdentifierName + "_"));
        }

        private static readonly DeclarationType[] ProcedureTypes =
        {
            DeclarationType.Procedure,
            DeclarationType.Function,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };

        private IEnumerable<Declaration> _interfaceMembers;

        /// <summary>
        /// Finds all interface members.
        /// </summary>
        public IEnumerable<Declaration> FindInterfaceMembers()
        {
            if (_interfaceMembers != null)
            {
                return _interfaceMembers;
            }

            var classes = _declarations.Where(item => item.DeclarationType == DeclarationType.Class);
            var interfaces = classes.Where(item => item.References.Any(reference =>
                reference.Context.Parent is VBAParser.ImplementsStmtContext))
                .Select(i => i.Scope)
                .ToList();

            _interfaceMembers = _declarations.Where(item => !item.IsBuiltIn 
                                                && ProcedureTypes.Contains(item.DeclarationType)
                                                && interfaces.Any(i => item.ParentScope.StartsWith(i)))
                                                .ToList();
            return _interfaceMembers;
        }

        private IEnumerable<Declaration> _interfaceImplementationMembers;

        /// <summary>
        /// Finds all class members that are interface implementation members.
        /// </summary>
        public IEnumerable<Declaration> FindInterfaceImplementationMembers()
        {
            if (_interfaceImplementationMembers != null)
            {
                return _interfaceImplementationMembers;
            }

            var members = FindInterfaceMembers();
            _interfaceImplementationMembers = _declarations.Where(item => !item.IsBuiltIn && ProcedureTypes.Contains(item.DeclarationType)
                && members.Select(m => m.ComponentName + '_' + m.IdentifierName).Contains(item.IdentifierName))
                .ToList();

            return _interfaceImplementationMembers;
        }

        public Declaration FindInterfaceMember(Declaration implementation)
        {
            var members = FindInterfaceMembers();
            var matches = members.Where(m => !m.IsBuiltIn && implementation.IdentifierName == m.ComponentName + '_' + m.IdentifierName).ToList();

            return matches.Count > 1 
                ? matches.SingleOrDefault(m => m.Project == implementation.Project) 
                : matches.First();
        }
    }
}
