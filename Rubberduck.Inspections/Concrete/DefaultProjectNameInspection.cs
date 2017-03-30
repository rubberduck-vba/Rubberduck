using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class DefaultProjectNameInspection : InspectionBase
    {
        private readonly IMessageBox _messageBox;

        public DefaultProjectNameInspection(RubberduckParserState state, IMessageBox messageBox)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
            _messageBox = messageBox;
        }

        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var projects = State.DeclarationFinder.UserDeclarations(DeclarationType.Project)
                .Where(item => item.IdentifierName.StartsWith("VBAProject"))
                .ToList();

            return projects
                .Select(issue => new DefaultProjectNameInspectionResult(this, issue, State, _messageBox))
                .ToList();
        }
    }
}
