using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class ImplicitActiveSheetReferenceInspection : InspectionBase
    {
        public ImplicitActiveSheetReferenceInspection(RubberduckParserState state)
            : base(state) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.MaintainabilityAndReadabilityIssues;

        private static readonly string[] Targets = 
        {
            "Cells", "Range", "Columns", "Rows"
        };

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var excel = State.DeclarationFinder.Projects.SingleOrDefault(item => !item.IsUserDefined && item.IdentifierName == "Excel");
            if (excel == null) { return Enumerable.Empty<IInspectionResult>(); }

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
                .Select(issue => new IdentifierReferenceInspectionResult(this,
                                                      string.Format(InspectionsUI.ImplicitActiveSheetReferenceInspectionResultFormat, issue.Declaration.IdentifierName),
                                                      State,
                                                      issue))
                .ToList();
        }
    }
}
