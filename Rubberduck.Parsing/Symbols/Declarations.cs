using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    public class Declarations
    {
        private readonly ConcurrentBag<Declaration> _declarations = new ConcurrentBag<Declaration>();

        public static readonly DeclarationType[] ProcedureTypes =
        {
            DeclarationType.Procedure,
            DeclarationType.Function,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };

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

        private IEnumerable<Declaration> _interfaces;
        private IEnumerable<Declaration> _interfaceMembers;

        /// <summary>
        /// Gets the <see cref="Declaration"/> of the specified <see cref="type"/>, 
        /// at the specified <see cref="selection"/>.
        /// Returns the declaration if selection is on an identifier reference.
        /// </summary>
        public Declaration FindSelectedDeclaration(QualifiedSelection selection, DeclarationType type, Func<Declaration, Selection> selector = null)
        {
            return FindSelectedDeclaration(selection, new[] {type}, selector);
        }

        public Declaration FindSelectedDeclaration(QualifiedSelection selection, IEnumerable<DeclarationType> types, Func<Declaration,Selection> selector = null)
        {
            var userDeclarations = _declarations.Where(item => !item.IsBuiltIn);
            var declarations = userDeclarations.Where(item => types.Contains(item.DeclarationType)
                && item.QualifiedName.QualifiedModuleName == selection.QualifiedName).ToList();

            var declaration = declarations.SingleOrDefault(item => 
                selector == null
                    ? item.Selection.Contains(selection.Selection)
                    : selector(item).Contains(selection.Selection));

            if (declaration != null)
            {
                return declaration;
            }

            // if we haven't returned yet, then we must be on an identifier reference.
            declaration = _declarations.SingleOrDefault(item => !item.IsBuiltIn
                && types.Contains(item.DeclarationType)
                && item.References.Any(reference =>
                reference.QualifiedModuleName == selection.QualifiedName
                && reference.Selection.Contains(selection.Selection)));

            return declaration;
        }

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

        /// <summary>
        /// Gets a tuple containing the <c>WithEvents</c> declaration and the corresponding handler,
        /// for each type implementing this event.
        /// </summary>
        public IEnumerable<Tuple<Declaration,Declaration>> FindHandlersForEvent(Declaration eventDeclaration)
        {
            return _declarations.Where(item => item.IsWithEvents && item.AsTypeName == eventDeclaration.ComponentName)
                .Select(item => new
                {
                    WithEventDeclaration = item, 
                    EventProvider = _declarations.SingleOrDefault(type => type.DeclarationType == DeclarationType.Class && type.QualifiedName.QualifiedModuleName == item.QualifiedName.QualifiedModuleName)
                })
                .Select(item => new
                {
                    WithEventsDeclaration = item.WithEventDeclaration,
                    ProviderEvents = _declarations.Where(member => member.DeclarationType == DeclarationType.Event && member.QualifiedSelection.QualifiedName == item.EventProvider.QualifiedName.QualifiedModuleName)
                })
                .Select(item => Tuple.Create(
                    item.WithEventsDeclaration,
                    _declarations.SingleOrDefault(declaration => declaration.DeclarationType == DeclarationType.Procedure
                    && declaration.QualifiedName.QualifiedModuleName == item.WithEventsDeclaration.QualifiedName.QualifiedModuleName
                    && declaration.IdentifierName == item.WithEventsDeclaration.IdentifierName + '_' + eventDeclaration.IdentifierName)
                    ));
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

        public Declaration FindSelection(QualifiedSelection selection, DeclarationType[] validDeclarationTypes)
        {
            var target = Items
                .Where(item => !item.IsBuiltIn)
                .FirstOrDefault(item => item.IsSelected(selection)
                                     || item.References.Any(r => r.IsSelected(selection)));

            if (target != null && validDeclarationTypes.Contains(target.DeclarationType))
            {
                return target;
            }

            target = null;

            var targets = Items
                .Where(item => !item.IsBuiltIn
                               && item.ComponentName == selection.QualifiedName.ComponentName
                               && validDeclarationTypes.Contains(item.DeclarationType));

            var currentSelection = new Selection(0, 0, int.MaxValue, int.MaxValue);

            foreach (var declaration in targets)
            {
                var activeSelection = new Selection(declaration.Context.Start.Line,
                                                    declaration.Context.Start.Column,
                                                    declaration.Context.Stop.Line,
                                                    declaration.Context.Stop.Column);

                if (currentSelection.Contains(activeSelection) && activeSelection.Contains(selection.Selection))
                {
                    target = declaration;
                    currentSelection = activeSelection;
                }

                foreach (var reference in declaration.References)
                {
                    var proc = (dynamic)reference.Context.Parent;
                    VBAParser.ArgsCallContext paramList;

                    // This is to prevent throws when this statement fails:
                    // (VBAParser.ArgsCallContext)proc.argsCall();
                    try { paramList = (VBAParser.ArgsCallContext)proc.argsCall(); }
                    catch { continue; }

                    if (paramList == null) { continue; }

                    activeSelection = new Selection(paramList.Start.Line,
                                                    paramList.Start.Column,
                                                    paramList.Stop.Line,
                                                    paramList.Stop.Column + paramList.Stop.Text.Length + 1);

                    if (currentSelection.Contains(activeSelection) && activeSelection.Contains(selection.Selection))
                    {
                        target = reference.Declaration;
                        currentSelection = activeSelection;
                    }
                }
            }
            return target;
        }
    }
}
