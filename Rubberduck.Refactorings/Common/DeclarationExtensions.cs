using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;

// ReSharper disable LocalizableElement

namespace Rubberduck.Common
{
    public static class DeclarationExtensions
    {
        /// <summary>
        /// Finds all event handler procedures for specified control declaration.
        /// </summary>
        public static IEnumerable<Declaration> FindEventHandlers(this IEnumerable<Declaration> declarations, Declaration control)
        {
            Debug.Assert(control.DeclarationType == DeclarationType.Control);

            return declarations.Where(declaration => declaration.ParentScope == control.ParentScope
                && declaration.DeclarationType == DeclarationType.Procedure
                && declaration.IdentifierName.StartsWith(control.IdentifierName + "_"));
        }

        public static IEnumerable<Declaration> FindUserEventHandlers(this IEnumerable<Declaration> declarations)
        {
            var declarationList = declarations.ToList();

            var userEvents =
                declarationList.Where(item => item.IsUserDefined && item.DeclarationType == DeclarationType.Event).ToList();

            var handlers = new List<Declaration>();
            foreach (var @event in userEvents)
            {
                handlers.AddRange(declarationList.FindHandlersForEvent(@event).Select(s => s.Item2));
            }
            
            return handlers;
        }

        public static IEnumerable<Declaration> FindFormEventHandlers(this RubberduckParserState state)
        {
            var items = state.AllDeclarations.ToList();

            var forms = items.Where(item => item.DeclarationType == DeclarationType.ClassModule
                && item.QualifiedName.QualifiedModuleName.ComponentType == ComponentType.UserForm)
                .ToList();

            var result = new List<Declaration>();
            foreach (var declaration in forms)
            {
                result.AddRange(FindFormEventHandlers(state, declaration));
            }

            return result;
        }

        public static IEnumerable<Declaration> FindFormEventHandlers(this RubberduckParserState state, Declaration userForm)
        {
            var items = state.AllDeclarations.ToList();
            var events = items.Where(item => !item.IsUserDefined
                                                     && item.ParentScope == "FM20.DLL;MSForms.FormEvents"
                                                     && item.DeclarationType == DeclarationType.Event).ToList();

            var handlerNames = events.Select(item => "UserForm_" + item.IdentifierName);
            var handlers = items.Where(item => item.ParentScope == userForm.Scope
                                                       && item.DeclarationType == DeclarationType.Procedure
                                                       && handlerNames.Contains(item.IdentifierName));

            return handlers.ToList();
        }

        /// <summary>
        /// Gets a tuple containing the <c>WithEvents</c> declaration and the corresponding handler,
        /// for each type implementing this event.
        /// </summary>
        public static IEnumerable<Tuple<Declaration,Declaration>> FindHandlersForEvent(this IEnumerable<Declaration> declarations, Declaration eventDeclaration)
        {
            var items = declarations as IList<Declaration> ?? declarations.ToList();
            return items.Where(item => item.IsWithEvents && item.AsTypeName == eventDeclaration.ComponentName)
            .Select(item => new
            {
                WithEventDeclaration = item, 
                EventProvider = items.SingleOrDefault(type => type.DeclarationType.HasFlag(DeclarationType.ClassModule) && type.QualifiedName.QualifiedModuleName == item.QualifiedName.QualifiedModuleName)
            })
            .Select(item => new
            {
                WithEventsDeclaration = item.WithEventDeclaration,
                ProviderEvents = items.Where(member => member.DeclarationType == DeclarationType.Event && member.QualifiedSelection.QualifiedName == item.EventProvider.QualifiedName.QualifiedModuleName)
            })
            .Select(item => Tuple.Create(
                item.WithEventsDeclaration,
                items.SingleOrDefault(declaration => declaration.DeclarationType == DeclarationType.Procedure
                && declaration.QualifiedName.QualifiedModuleName == item.WithEventsDeclaration.QualifiedName.QualifiedModuleName
                && declaration.IdentifierName == item.WithEventsDeclaration.IdentifierName + '_' + eventDeclaration.IdentifierName)
                ));
        }

        public static IEnumerable<Declaration> FindEventProcedures(this IEnumerable<Declaration> declarations, Declaration withEventsDeclaration)
        {
            if (!withEventsDeclaration.IsWithEvents)
            {
                return new Declaration[]{};
            }

            var items = declarations as IList<Declaration> ?? declarations.ToList();
            var type = withEventsDeclaration.AsTypeDeclaration;

            if (type == null)
            {
                return new Declaration[]{};
            }

            var members = items
                .Where(item => Equals(item.ParentScopeDeclaration, type))
                .ToList();
            var events = members.Where(member => member.DeclarationType == DeclarationType.Event);
            var handlerNames = events.Select(e => withEventsDeclaration.IdentifierName + '_' + e.IdentifierName);

            return items.Where(item => item.Project != null 
                                               && item.ProjectId == withEventsDeclaration.ProjectId
                                               && item.ParentScope == withEventsDeclaration.ParentScope
                                               && item.DeclarationType == DeclarationType.Procedure
                                               && handlerNames.Any(name => item.IdentifierName == name))
                .ToList();
        }
    }
}
