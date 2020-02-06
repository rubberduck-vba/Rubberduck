using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Settings;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Inspections.Results;
using Rubberduck.JunkDrawer.Extensions;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Refactorings.Common;
using Rubberduck.SettingsProvider;
using Rubberduck.VBEditor;
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
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub CpFrmtRls(ByVal rng1 As Range, ByVal rng2 As Range)
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub CopyFormatRules(ByVal source As Range, ByVal destination As Range)
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class UseMeaningfulNameInspection : InspectionBase
    {
        private readonly IConfigurationService<CodeInspectionSettings> _settings;

        public UseMeaningfulNameInspection(RubberduckParserState state, IConfigurationService<CodeInspectionSettings> settings)
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
            var finder = DeclarationFinderProvider.DeclarationFinder;
            var settings = _settings.Read();
            var whitelistedNames = settings.WhitelistedIdentifiers
                .Select(s => s.Identifier)
                .ToArray();
            var handlers = finder.FindEventHandlers().ToHashSet();

            var results = new List<IInspectionResult>();
            foreach (var moduleDeclaration in State.DeclarationFinder.UserDeclarations(DeclarationType.Module))
            {
                if (moduleDeclaration == null)
                {
                    continue;
                }

                var module = moduleDeclaration.QualifiedModuleName;
                results.AddRange(DoGetInspectionResults(module, finder, whitelistedNames, handlers));
            }

            return results;
        }

        private IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module)
        {
            var finder = DeclarationFinderProvider.DeclarationFinder;
            var settings = _settings.Read();
            var whitelistedNames = settings.WhitelistedIdentifiers
                .Select(s => s.Identifier)
                .ToArray();
            var handlers = finder.FindEventHandlers().ToHashSet();
            return DoGetInspectionResults(module, finder, whitelistedNames, handlers);
        }

        private IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder, string[] whitelistedNames, ICollection<Declaration> eventHandlers)
        {
            var objectionableDeclarations = finder.Members(module)
                .Where(declaration => IsResultDeclaration(declaration, whitelistedNames, eventHandlers));

            return objectionableDeclarations
                .Select(InspectionResult)
                .ToList();
        }

        private static bool IsResultDeclaration(Declaration declaration, string[] whitelistedNames, ICollection<Declaration> eventHandlers)
        {
            return !string.IsNullOrEmpty(declaration.IdentifierName)
                   && !IgnoreDeclarationTypes.Contains(declaration.DeclarationType)
                   && !(declaration.Context is LineNumberLabelContext)
                   && (declaration.ParentDeclaration == null
                       || !IgnoreDeclarationTypes.Contains(declaration.ParentDeclaration.DeclarationType)
                            && !eventHandlers.Contains(declaration.ParentDeclaration))
                   && !whitelistedNames.Contains(declaration.IdentifierName)
                   && !VBAIdentifierValidator.IsMeaningfulIdentifier(declaration.IdentifierName);
        }

        private IInspectionResult InspectionResult(Declaration declaration)
        {
            return new DeclarationInspectionResult(
                this,
                ResultDescription(declaration),
                declaration,
                disabledQuickFixes: DisabledQuickFixes(declaration));
        }

        private static string ResultDescription(Declaration declaration)
        {
            var declarationType = declaration.DeclarationType.ToLocalizedString();
            var declarationName = declaration.IdentifierName;
            return string.Format(
                Resources.Inspections.InspectionResults.IdentifierNameInspection,
                declarationType,
                declarationName);
        }

        private static ICollection<string> DisabledQuickFixes(Declaration declaration)
        {
            return declaration.DeclarationType.HasFlag(DeclarationType.Module)
                   || declaration.DeclarationType.HasFlag(DeclarationType.Project)
                   ? new List<string> {"IgnoreOnceQuickFix"}
                   : new List<string>();
        }
    }
}
