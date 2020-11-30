using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    internal sealed class PublicEnumerationDeclaredWithinWorksheetInspection : DeclarationInspectionBase
    {
        public PublicEnumerationDeclaredWithinWorksheetInspection(IDeclarationFinderProvider declarationFinderProvider, IProjectsProvider projectsProvider)
            : base(declarationFinderProvider, new DeclarationType[] { DeclarationType.Enumeration })
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            return declaration.Accessibility == Accessibility.Public
                && declaration.QualifiedModuleName.ComponentType == ComponentType.Document;
        }

        protected override string ResultDescription(Declaration declaration)
        {
            return string.Format(InspectionResults.PublicEnumerationDeclaredWithinWorksheetInspection,
                declaration.IdentifierName,
                declaration.ParentScopeDeclaration.IdentifierName);
        }
    }
}
