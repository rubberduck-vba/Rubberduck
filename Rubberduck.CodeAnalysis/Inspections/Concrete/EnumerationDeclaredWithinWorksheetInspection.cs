using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    internal class PublicEnumerationDeclaredWithinWorksheetInspection : DeclarationInspectionBase
    {
        private readonly IProjectsProvider _projectsProvider;

        public PublicEnumerationDeclaredWithinWorksheetInspection(IDeclarationFinderProvider declarationFinderProvider, IProjectsProvider projectsProvider)
            : base(declarationFinderProvider)
        {
            _projectsProvider = projectsProvider;
        }

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            return declaration.Accessibility == Accessibility.Public
                && declaration.DeclarationType == DeclarationType.Enumeration
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
