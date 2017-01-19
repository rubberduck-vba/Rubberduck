using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public sealed class UseMeaningfulNameInspection : InspectionBase
    {
        private readonly IMessageBox _messageBox;
        private readonly IPersistanceService<CodeInspectionSettings> _settings;

        public UseMeaningfulNameInspection(IMessageBox messageBox, RubberduckParserState state, IPersistanceService<CodeInspectionSettings> settings)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
            _messageBox = messageBox;
            _settings = settings;
        }

        public override string Description { get { return InspectionsUI.UseMeaningfulNameInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var settings = _settings.Load(new CodeInspectionSettings()) ?? new CodeInspectionSettings();
            var whitelistedNames = settings.WhitelistedIdentifiers.Select(s => s.Identifier).ToList();

            var issues = UserDeclarations
                            .Where(declaration => declaration.DeclarationType != DeclarationType.ModuleOption &&
                                                  !whitelistedNames.Contains(declaration.IdentifierName) &&
                                                  (declaration.IdentifierName.Length < 3 ||
                                                  char.IsDigit(declaration.IdentifierName.Last()) ||
                                                  !declaration.IdentifierName.Any(c => 
                                                      "aeiouy".Any(a => string.Compare(a.ToString(), c.ToString(), StringComparison.OrdinalIgnoreCase) == 0))))
                            .Select(issue => new IdentifierNameInspectionResult(this, issue, State, _messageBox, _settings))
                            .ToList();

            return issues;
        }
    }
}
