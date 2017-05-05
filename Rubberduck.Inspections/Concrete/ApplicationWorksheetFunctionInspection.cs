using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete
{
    public class ApplicationWorksheetFunctionInspection : InspectionBase
    {
        public ApplicationWorksheetFunctionInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var excel = State.DeclarationFinder.Projects.SingleOrDefault(item => !item.IsUserDefined && item.IdentifierName == "Excel");
            if (excel == null) { return Enumerable.Empty<IInspectionResult>(); }

            var members = new HashSet<string>(BuiltInDeclarations.Where(decl => decl.DeclarationType == DeclarationType.Function &&
                                                                        decl.ParentDeclaration != null && 
                                                                        decl.ParentDeclaration.ComponentName.Equals("WorksheetFunction"))
                                                                 .Select(decl => decl.IdentifierName));

            var usages = BuiltInDeclarations.Where(decl => decl.References.Any() &&
                                                           decl.ProjectName.Equals("Excel") &&
                                                           decl.ComponentName.Equals("Application") &&
                                                           members.Contains(decl.IdentifierName));

            return from usage in usages
                   from reference in usage.References.Where(use => !IsIgnoringInspectionResultFor(use, AnnotationName))
                   let qualifiedSelection = new QualifiedSelection(reference.QualifiedModuleName, reference.Selection)
                   select new IdentifierReferenceInspectionResult(this,
                                               string.Format(InspectionsUI.ApplicationWorksheetFunctionInspectionResultFormat, usage.IdentifierName),
                                               State,
                                               reference);
        }
    }
}
