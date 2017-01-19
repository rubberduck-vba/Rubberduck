using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class ImplicitActiveWorkbookReferenceInspection : InspectionBase
    {
        public ImplicitActiveWorkbookReferenceInspection(RubberduckParserState state)
            : base(state)
        {
        }

        public override string Meta { get { return InspectionsUI.ImplicitActiveWorkbookReferenceInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ImplicitActiveWorkbookReferenceInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }

        private static readonly string[] Targets =
        {
            "Worksheets", "Sheets", "Names", "_Default"
        };

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var excel = State.DeclarationFinder.Projects.SingleOrDefault(item => item.IsBuiltIn && item.IdentifierName == "Excel");
            if (excel == null) { return Enumerable.Empty<InspectionResultBase>(); }

            var modules = new[]
            {
                State.DeclarationFinder.FindClassModule("_Global", excel, true),
                State.DeclarationFinder.FindClassModule("_Application", excel, true),
                State.DeclarationFinder.FindClassModule("Global", excel, true),
                State.DeclarationFinder.FindClassModule("Application", excel, true),
                State.DeclarationFinder.FindClassModule("Sheets", excel, true),
            };

            var members = Targets
                .SelectMany(target => modules.SelectMany(module =>
                    State.DeclarationFinder.FindMemberMatches(module, target)))
                .Where(item => item.References.Any())
                .SelectMany(item => item.References.Where(reference => !IsIgnoringInspectionResultFor(reference, AnnotationName)))
                .ToList();
                
            return members.Select(issue => new ImplicitActiveWorkbookReferenceInspectionResult(this, issue));
        }
    }
}
