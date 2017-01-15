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

        private static readonly string[] ParentScopes =
        {
            "_Global",
            "Global",
            "_Application",
            "Application",
            "Sheets",
            //"Worksheets",
        };

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var issues = BuiltInDeclarations
                .Where(item => item.ProjectName == "Excel" && ParentScopes.Contains(item.ComponentName) 
                    && item.References.Any(r => Targets.Contains(r.IdentifierName)))
                .SelectMany(declaration => declaration.References.Distinct())
                .ToList();
                
            var filtered = issues.Where(item => Targets.Contains(item.IdentifierName));
            return filtered.Select(issue => new ImplicitActiveWorkbookReferenceInspectionResult(this, issue));
        }
    }
}
