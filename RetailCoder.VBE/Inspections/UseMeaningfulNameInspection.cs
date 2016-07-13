using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public sealed class UseMeaningfulNameInspection : InspectionBase
    {
        private readonly IMessageBox _messageBox;

        public UseMeaningfulNameInspection(IMessageBox messageBox, RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
            _messageBox = messageBox;
        }

        public override string Description { get { return InspectionsUI.UseMeaningfulNameInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var issues = UserDeclarations
                            .Where(declaration => declaration.DeclarationType != DeclarationType.ModuleOption && 
                                                  (declaration.IdentifierName.Length < 3 ||
                                                  char.IsDigit(declaration.IdentifierName.Last()) ||
                                                  !declaration.IdentifierName.Any(c => 
                                                      "aeiouy".Any(a => string.Compare(a.ToString(), c.ToString(), StringComparison.OrdinalIgnoreCase) == 0))))
                            .Select(issue => new UseMeaningfulNameInspectionResult(this, issue, State, _messageBox))
                            .ToList();

            return issues;
        }
    }
}
