using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class ImplicitActiveSheetReferenceInspection : InspectionBase
    {
        public ImplicitActiveSheetReferenceInspection(RubberduckParserState state)
            : base(state)
        {
        }

        public override string Meta { get { return InspectionsUI.ImplicitActiveSheetReferenceInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ImplicitActiveSheetReferenceInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }

        private static readonly string[] Targets = 
        {
            "Cells", "Range", "Columns", "Rows"
        };

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var matches = BuiltInDeclarations.Where(item =>
                        item.ProjectName == "Excel" &&
                        Targets.Contains(item.IdentifierName) &&
                        (item.ParentDeclaration.ComponentName == "_Global" || item.ParentDeclaration.ComponentName == "Global") &&
                        item.AsTypeName == "Range").ToList();

            var issues = matches.Where(item => item.References.Any())
                .SelectMany(declaration => declaration.References.Distinct());

            return issues
                .Where(issue => !issue.IsInspectionDisabled(AnnotationName))
                .Select(issue => new ImplicitActiveSheetReferenceInspectionResult(this, issue));
        }
    }
}
