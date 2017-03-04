using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class ApplicationWorksheetFunctionInspection : InspectionBase
    {
        public ApplicationWorksheetFunctionInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        { }

        public override string Meta { get { return InspectionsUI.ApplicationWorksheetFunctionInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ApplicationWorksheetFunctionInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var excel = State.DeclarationFinder.Projects.SingleOrDefault(item => item.IsBuiltIn && item.IdentifierName == "Excel");
            if (excel == null) { return Enumerable.Empty<InspectionResultBase>(); }

            var members = new HashSet<string>(BuiltInDeclarations.Where(decl => decl.DeclarationType == DeclarationType.Function &&
                                                                        decl.ParentDeclaration != null && 
                                                                        decl.ParentDeclaration.ComponentName.Equals("WorksheetFunction"))
                                                                 .Select(decl => decl.IdentifierName));

            var usages = BuiltInDeclarations.Where(decl => decl.References.Any() &&
                                                           decl.ProjectName.Equals("Excel") &&
                                                           decl.ComponentName.Equals("Application") &&
                                                           members.Contains(decl.IdentifierName));

            return (from usage in usages
                from reference in usage.References.Where(use => !IsIgnoringInspectionResultFor(use, AnnotationName))
                let qualifiedSelection = new QualifiedSelection(reference.QualifiedModuleName, reference.Selection)
                select new ApplicationWorksheetFunctionInspectionResult(this, qualifiedSelection, usage.IdentifierName));
        }
    }
}
