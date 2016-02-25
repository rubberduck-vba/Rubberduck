using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Inspections
{
    public sealed class UseMeaningfulNameInspection : InspectionBase
    {
        private readonly IMessageBox _messageBox;
        private readonly ICodePaneWrapperFactory _wrapperFactory;

        public UseMeaningfulNameInspection(IMessageBox messageBox, RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
            _messageBox = messageBox;
            _wrapperFactory = new CodePaneWrapperFactory();
        }

        public override string Description { get { return InspectionsUI.UseMeaningfulNameInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var issues = UserDeclarations
                            .Where(declaration => declaration.DeclarationType != DeclarationType.ModuleOption &&
                                                  (declaration.IdentifierName.Length < 3 ||
                                                  char.IsDigit(declaration.IdentifierName.Last()) ||
                                                  !declaration.IdentifierName.Any(c => new[] {'a', 'e', 'i', 'o', 'u', 'y'}.Contains(c))))
                            .Select(issue => new UseMeaningfulNameInspectionResult(this, issue, State, _wrapperFactory, _messageBox))
                            .ToList();

            return issues;
        }
    }
}