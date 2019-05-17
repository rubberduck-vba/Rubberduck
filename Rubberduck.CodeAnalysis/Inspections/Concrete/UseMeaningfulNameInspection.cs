using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Rubberduck.CodeAnalysis.Settings;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources;
using Rubberduck.SettingsProvider;
using static Rubberduck.Parsing.Grammar.VBAParser;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Warns about identifiers that have names that are likely to be too short, disemvoweled, or appended with a numeric suffix.
    /// </summary>
    /// <why>
    /// Meaningful, pronounceable, unabbreviated names read better and leave less room for interpretation. 
    /// Moreover, names suffixed with a number can indicate the need to look into an array, collection, or dictionary data structure.
    /// </why>
    /// <example>
    /// <![CDATA[
    /// Public Sub CpFrmtRls(ByVal rng1 As Range, ByVal rng2 As Range)
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    /// <example>
    /// <![CDATA[
    /// Public Sub CopyFormatRules(ByVal source As Range, ByVal destination As Range)
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class UseMeaningfulNameInspection : InspectionBase
    {
        private readonly IPersistenceService<CodeInspectionSettings> _settings;

        public UseMeaningfulNameInspection(RubberduckParserState state, IPersistenceService<CodeInspectionSettings> settings)
            : base(state)
        {
            _settings = settings;
        }

        private static readonly DeclarationType[] IgnoreDeclarationTypes = 
        {
            DeclarationType.BracketedExpression, 
            DeclarationType.LibraryFunction,
            DeclarationType.LibraryProcedure, 
        };

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var settings = _settings.Load(new CodeInspectionSettings()) ?? new CodeInspectionSettings();
            var whitelistedNames = settings.WhitelistedIdentifiers.Select(s => s.Identifier).ToArray();

            var handlers = State.DeclarationFinder.FindEventHandlers();

            var issues = UserDeclarations
                            .Where(declaration => !string.IsNullOrEmpty(declaration.IdentifierName) &&
                                !IgnoreDeclarationTypes.Contains(declaration.DeclarationType) &&
                                !(declaration.Context is LineNumberLabelContext) &&
                                (declaration.ParentDeclaration == null || 
                                    !IgnoreDeclarationTypes.Contains(declaration.ParentDeclaration.DeclarationType) &&
                                    !handlers.Contains(declaration.ParentDeclaration)) &&
                                !whitelistedNames.Contains(declaration.IdentifierName) &&
                                !VariableNameValidator.IsMeaningfulName(declaration.IdentifierName));

            return (from issue in issues select CreateInspectionResult(this, issue))
                .ToList();
        }

        private static DeclarationInspectionResult CreateInspectionResult(IInspection inspection, Declaration issue)
        {
            dynamic properties = null;

            if (issue.DeclarationType.HasFlag(DeclarationType.Module) ||
                issue.DeclarationType.HasFlag(DeclarationType.Project))
            {
                properties = new PropertyBag();
                properties.DisableFixes = "IgnoreOnceQuickFix";
            }

            return new DeclarationInspectionResult(inspection,
                string.Format(Resources.Inspections.InspectionResults.IdentifierNameInspection,
                    RubberduckUI.ResourceManager.GetString("DeclarationType_" + issue.DeclarationType,
                        CultureInfo.CurrentUICulture), issue.IdentifierName),
                issue, properties: properties);
        }
    }
}
