using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class DefaultProjectNameInspection : InspectionBase
    {
        public DefaultProjectNameInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.MaintainabilityAndReadabilityIssues;

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var projects = State.DeclarationFinder.UserDeclarations(DeclarationType.Project)
                .Where(item => item.IdentifierName.StartsWith("VBAProject"))
                .ToList();

            return projects
                .Select(issue => new DeclarationInspectionResult(this, Description, issue))
                .ToList();
        }
    }
}
