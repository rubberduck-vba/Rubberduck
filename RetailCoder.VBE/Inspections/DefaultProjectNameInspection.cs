using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class DefaultProjectNameInspection : InspectionBase
    {
        public DefaultProjectNameInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
        }

        public override string Meta { get { return InspectionsUI.DefaultProjectNameInspectionMeta; } }
        public override string Description { get { return InspectionsUI.DefaultProjectNameInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var projects = State.DeclarationFinder.UserDeclarations(DeclarationType.Project)
                .Where(item => item.IdentifierName.StartsWith("VBAProject"))
                .ToList();

            return projects
                .Select(issue => new DefaultProjectNameInspectionResult(this, issue, State))
                .ToList();
        }
    }
}
