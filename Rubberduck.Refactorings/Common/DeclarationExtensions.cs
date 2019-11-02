using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Common
{
    public static class DeclarationExtensions
    {
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
    }
}
