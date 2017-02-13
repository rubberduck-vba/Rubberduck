using System;
using System.Collections.Generic;
using System.Linq;
using Castle.Core.Internal;
using Rubberduck.Common;
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

        private static readonly DeclarationType[] IgnoreDeclarationTypes = 
        {
            DeclarationType.ModuleOption,
            DeclarationType.BracketedExpression, 
            DeclarationType.LibraryFunction,
            DeclarationType.LibraryProcedure, 
        };

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var settings = _settings.Load(new CodeInspectionSettings()) ?? new CodeInspectionSettings();
            var whitelistedNames = settings.WhitelistedIdentifiers.Select(s => s.Identifier).ToArray();

            var handlers = State.DeclarationFinder.FindBuiltinEventHandlers();

            var issues = UserDeclarations
                            .Where(declaration => !string.IsNullOrEmpty(declaration.IdentifierName) &&
                                !IgnoreDeclarationTypes.Contains(declaration.DeclarationType) &&
                                (declaration.ParentDeclaration == null || 
                                    !IgnoreDeclarationTypes.Contains(declaration.ParentDeclaration.DeclarationType) &&
                                    !handlers.Contains(declaration.ParentDeclaration)) &&
                                !whitelistedNames.Contains(declaration.IdentifierName) &&
                                IsBadIdentifier(declaration.IdentifierName))
                            .Select(issue => new IdentifierNameInspectionResult(this, issue, State, _messageBox, _settings))
                            .ToList();

            return issues;
        }

        private static bool IsBadIdentifier(string identifier)
        {
            var validator = new VariableNameValidator(identifier);
            return !validator.IsMeaningfulName();
        }
    }
}
