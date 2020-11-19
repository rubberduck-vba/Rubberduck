using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Antlr4.Runtime.Misc;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    internal class EnumerationDeclaredWithinWorksheetInspection : InspectionBase
    {
        private readonly IProjectsProvider _projectsProvider;

        public EnumerationDeclaredWithinWorksheetInspection(IDeclarationFinderProvider declarationFinderProvider, IProjectsProvider projectsProvider)
            : base(declarationFinderProvider)
        {
            _projectsProvider = projectsProvider;
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(DeclarationFinder finder)
        {
            var declaredEnums = finder.DeclarationsWithType(DeclarationType.Enumeration)
                .Where(d => d.QualifiedModuleName.ComponentType == VBEditor.SafeComWrappers.ComponentType.DocObject);
            
            return declaredEnums.Select(d => InspectionResult(d));
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder)
        {
            var declaredEnums = finder.DeclarationsWithType(DeclarationType.Enumeration)
                .Where(d => d.QualifiedModuleName == module);

            return declaredEnums.Select(d => InspectionResult(d));
        }

        protected virtual IInspectionResult InspectionResult(Declaration declaration)
        {
            return new DeclarationInspectionResult(this, InspectionResults.EnumerationDeclaredWithinWorksheetInspection, declaration);
        }
    }
}
