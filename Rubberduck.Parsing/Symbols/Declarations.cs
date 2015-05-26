using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
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

        private IEnumerable<Declaration> _interfaces;
        private IEnumerable<Declaration> _interfaceMembers;

        public IEnumerable<Declaration> FindInterfaces()
        {
            if (_interfaces != null)
            {
                return _interfaces;
            }

            var classes = _declarations.Where(item => item.DeclarationType == DeclarationType.Class);
            _interfaces = classes.Where(item => item.References.Any(reference =>
                reference.Context.Parent is VBAParser.ImplementsStmtContext))
                .ToList();

            return _interfaces;
        }

        /// <summary>
        /// Finds all interface members.
        /// </summary>
        public IEnumerable<Declaration> FindInterfaceMembers()
        {
            if (_interfaceMembers != null)
            {
                return _interfaceMembers;
            }

            var interfaces = FindInterfaces().Select(i => i.Scope).ToList();
            _interfaceMembers = _declarations.Where(item => !item.IsBuiltIn 
                                                && ProcedureTypes.Contains(item.DeclarationType)
                                                && interfaces.Any(i => item.ParentScope.StartsWith(i)))
                                                .ToList();
            return _interfaceMembers;
        }

        public IEnumerable<Declaration> FindFormEventHandlers()
        {
            var forms = _declarations.Where(item => item.DeclarationType == DeclarationType.Class
                && item.QualifiedName.QualifiedModuleName.Component != null
                && item.QualifiedName.QualifiedModuleName.Component.Type == vbext_ComponentType.vbext_ct_MSForm)
                .ToList();

            var result = new List<Declaration>();
            foreach (var declaration in forms)
            {
                result.AddRange(FindFormEventHandlers(declaration));
            }

            return result;
        }

        public IEnumerable<Declaration> FindFormEventHandlers(Declaration userForm)
        {
            var events = _declarations.Where(item => item.IsBuiltIn
                                                     && item.ParentScope == "MSForms.UserForm"
                                                     && item.DeclarationType == DeclarationType.Event).ToList();
            var handlerNames = events.Select(item => "UserForm_" + item.IdentifierName);
            var handlers = _declarations.Where(item => item.ParentScope == userForm.Scope
                                                       && item.DeclarationType == DeclarationType.Procedure
                                                       && handlerNames.Contains(item.IdentifierName));

            return handlers.ToList();
        }

        public IEnumerable<Declaration> FindEventProcedures(Declaration withEventsDeclaration)
        {
            if (!withEventsDeclaration.IsWithEvents)
            {
                return new Declaration[]{};
            }

            var type = _declarations.SingleOrDefault(item => item.DeclarationType == DeclarationType.Class
                                                             && item.Project != null
                                                             && item.IdentifierName == withEventsDeclaration.AsTypeName.Split('.').Last());

            if (type == null)
            {
                return new Declaration[]{};
            }

            var members = GetTypeMembers(type).ToList();
            var events = members.Where(member => member.DeclarationType == DeclarationType.Event);
            var handlerNames = events.Select(e => withEventsDeclaration.IdentifierName + '_' + e.IdentifierName);

            return _declarations.Where(item => item.Project != null 
                                               && item.Project.Equals(withEventsDeclaration.Project)
                                               && item.ParentScope == withEventsDeclaration.ParentScope
                                               && item.DeclarationType == DeclarationType.Procedure
                                               && handlerNames.Any(name => item.IdentifierName == name))
                .ToList();
        }

        private IEnumerable<Declaration> GetTypeMembers(Declaration type)
        {
            return _declarations.Where(item => item.Project != null && item.Project.Equals(type.Project) && item.ParentScope == type.Scope);
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

        public IEnumerable<Declaration> FindInterfaceImplementationMembers(string interfaceMember)
        {
            return FindInterfaceImplementationMembers()
                .Where(m => m.IdentifierName.EndsWith(interfaceMember));
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
