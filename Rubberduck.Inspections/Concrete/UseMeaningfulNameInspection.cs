using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class UseMeaningfulNameInspection : InspectionBase
    {
        private readonly IPersistanceService<CodeInspectionSettings> _settings;

        public UseMeaningfulNameInspection(RubberduckParserState state, IPersistanceService<CodeInspectionSettings> settings)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
            _settings = settings;
        }

        public override CodeInspectionType InspectionType => CodeInspectionType.MaintainabilityAndReadabilityIssues;

        private static readonly DeclarationType[] IgnoreDeclarationTypes = 
        {
            DeclarationType.ModuleOption,
            DeclarationType.BracketedExpression, 
            DeclarationType.LibraryFunction,
            DeclarationType.LibraryProcedure, 
        };

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var settings = _settings.Load(new CodeInspectionSettings()) ?? new CodeInspectionSettings();
            var whitelistedNames = settings.WhitelistedIdentifiers.Select(s => s.Identifier).ToArray();

            var handlers = State.DeclarationFinder.FindEventHandlers();

            var issues = UserDeclarations
                            .Where(declaration => !string.IsNullOrEmpty(declaration.IdentifierName) &&
                                !IgnoreDeclarationTypes.Contains(declaration.DeclarationType) &&
                                (declaration.ParentDeclaration == null || 
                                    !IgnoreDeclarationTypes.Contains(declaration.ParentDeclaration.DeclarationType) &&
                                    !handlers.Contains(declaration.ParentDeclaration)) &&
                                !whitelistedNames.Contains(declaration.IdentifierName) &&
                                !VariableNameValidator.IsMeaningfulName(declaration.IdentifierName))
                            .Select(issue => new DeclarationInspectionResult(this,
                                                                  string.Format(InspectionsUI.IdentifierNameInspectionResultFormat,
                                                                                RubberduckUI.ResourceManager.GetString("DeclarationType_" + issue.DeclarationType, CultureInfo.CurrentUICulture),
                                                                                issue.IdentifierName),
                                                                  issue))
                            .ToList();

            return issues;
        }
    }
}
