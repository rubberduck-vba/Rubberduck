using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    internal class EnumerationDeclaredWithinWorksheetInspection : DeclarationInspectionBase
    {
        private readonly IProjectsProvider _projectsProvider;

        public EnumerationDeclaredWithinWorksheetInspection(IDeclarationFinderProvider declarationFinderProvider, IProjectsProvider projectsProvider)
            : base(declarationFinderProvider)
        {
            _projectsProvider = projectsProvider;
        }

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            return declaration.DeclarationType == DeclarationType.Enumeration
                && declaration.QualifiedModuleName.ComponentType == ComponentType.DocObject;
        }

        protected override string ResultDescription(Declaration declaration)
        {
            return string.Format(InspectionResults.EnumerationDeclaredWithinWorksheetInspection,
                declaration.IdentifierName,
                declaration.ParentScopeDeclaration.IdentifierName);
        }
    }
}
