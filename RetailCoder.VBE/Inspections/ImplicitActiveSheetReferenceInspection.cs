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
            var excel = State.DeclarationFinder.Projects.SingleOrDefault(item => item.IsBuiltIn && item.IdentifierName == "Excel");
            if (excel == null) { return Enumerable.Empty<InspectionResultBase>(); }

            var globalModules = new[]
            {
                State.DeclarationFinder.FindClassModule("Global", excel, true),
                State.DeclarationFinder.FindClassModule("_Global", excel, true)
            };

            var members = Targets
                .SelectMany(target => globalModules.SelectMany(global =>
                    State.DeclarationFinder.FindMemberMatches(global, target))
                .Where(member => member.AsTypeName == "Range" && member.References.Any()));

            return members
                .SelectMany(declaration => declaration.References)
                .Where(issue => !issue.IsIgnoringInspectionResultFor(AnnotationName))
                .Select(issue => new ImplicitActiveSheetReferenceInspectionResult(this, issue))
                .ToList();
        }
    }
}
