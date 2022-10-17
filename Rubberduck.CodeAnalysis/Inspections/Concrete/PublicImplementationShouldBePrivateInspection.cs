using Rubberduck.CodeAnalysis.CodeMetrics;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Refactorings.Common;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    internal sealed class PublicImplementationShouldBePrivateInspection : DeclarationInspectionBase
    {
        public PublicImplementationShouldBePrivateInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, DeclarationType.Member)
        {}

        //Overriding DoGetInspectionResults in order to dereference the DeclarationFinder FindXXX declaration 
        //lists only once per inspections pass.
        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(DeclarationFinder finder)
        {
            var publicMembers = finder.UserDeclarations(DeclarationType.Member)
                .Where(d => !d.HasPrivateAccessibility()
                    && IsLikeAnImplementerOrHandlerName(d.IdentifierName));

            if (!publicMembers.Any())
            {
                return Enumerable.Empty<IInspectionResult>();
            }

            var publicImplementersAndHandlers = finder.FindAllInterfaceImplementingMembers()
                .Where(d => !d.HasPrivateAccessibility())
                .Concat(finder.FindEventHandlers()
                    .Where(d => !d.HasPrivateAccessibility()));

            var publicDocumentEvents = FindDocumentEventHandlers(publicMembers);

            return publicMembers.Intersect(publicImplementersAndHandlers)
                .Concat(publicDocumentEvents)
                .Select(InspectionResult)
                .ToList();
        }

        private static IEnumerable<Declaration> FindDocumentEventHandlers(IEnumerable<Declaration> publicMembers)
        {
            //Excel and Word
            var docEventPrefixes = new List<string>() 
            { 
                "Workbook", 
                "Worksheet", 
                "Document" 
            };

            //FindDocumentEventHandlers can be a source of False Positives if a Document's code
            //contains Public procedure Identifiers (with a single underscore).
            return publicMembers.Where(d => d.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Document)
                && d.DeclarationType.Equals(DeclarationType.Procedure)
                && docEventPrefixes.Any(dep => IsLikeADocumentEventHandlerName(d.IdentifierName, dep)));
        }

        protected override string ResultDescription(Declaration declaration)
        {
            return string.Format(Resources.Inspections.InspectionResults.PublicImplementationShouldBePrivateInspection, 
                declaration.IdentifierName);
        }

        private static bool IsLikeAnImplementerOrHandlerName(string identifier)
        {
            var splitup = identifier.Split('_');
            return splitup.Length == 2 && splitup[1].Length > 0;
        }

        private static bool IsLikeADocumentEventHandlerName(string procedureName, string docEventHandlerPrefix)
        {
            var splitup = procedureName.Split('_');

            return splitup.Length == 2 
                && splitup[0].Equals(docEventHandlerPrefix, StringComparison.InvariantCultureIgnoreCase)
                && splitup[1].Length > 2 //Excel and Word document events all have at least 3 characters
                && !splitup[1].Any(c => char.IsDigit(c)); //Excel and Word document event names do not contain numbers
        }

        //The 'DoGetInspectionResults' override excludes IsResultDeclaration from the execution path
        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            throw new NotImplementedException();
        }

    }
}
