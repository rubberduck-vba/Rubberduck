using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Extensions;
using Rubberduck.CodeAnalysis.Settings;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Refactorings.Common;
using Rubberduck.SettingsProvider;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Warns about identifiers that have names that are likely to be too short, disemvoweled, or appended with a numeric suffix.
    /// </summary>
    /// <why>
    /// Meaningful, pronounceable, unabbreviated names read better and leave less room for interpretation. 
    /// Moreover, names suffixed with a number can indicate the need to look into an array, collection, or dictionary data structure.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub CpFrmtRls(ByVal rng1 As Range, ByVal rng2 As Range)
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub CopyFormatRules(ByVal source As Range, ByVal destination As Range)
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class UseMeaningfulNameInspection : DeclarationInspectionUsingGlobalInformationBase<string[]>
    {
        private readonly IConfigurationService<CodeInspectionSettings> _settings;

        public UseMeaningfulNameInspection(IDeclarationFinderProvider declarationFinderProvider, IConfigurationService<CodeInspectionSettings> settings)
            : base(declarationFinderProvider)
        {
            _settings = settings;
        }

        private static readonly DeclarationType[] IgnoreDeclarationTypes = 
        {
            DeclarationType.BracketedExpression, 
            DeclarationType.LibraryFunction,
            DeclarationType.LibraryProcedure, 
        };

        protected override string[] GlobalInformation(DeclarationFinder finder)
        {
            var settings = _settings.Read();
            return settings.WhitelistedIdentifiers
                .Select(s => s.Identifier)
                .ToArray();
        }

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder, string[] whitelistedNames)
        {
            return !string.IsNullOrEmpty(declaration.IdentifierName)
                   && !IgnoreDeclarationTypes.Contains(declaration.DeclarationType)
                   && !(declaration.Context is VBAParser.LineNumberLabelContext)
                   && (declaration.ParentDeclaration == null
                       || !IgnoreDeclarationTypes.Contains(declaration.ParentDeclaration.DeclarationType)
                       && !finder.FindEventHandlers().Contains(declaration.ParentDeclaration))
                   && !whitelistedNames.Contains(declaration.IdentifierName)
                   && !VBAIdentifierValidator.IsMeaningfulIdentifier(declaration.IdentifierName);
        }

        protected override string ResultDescription(Declaration declaration)
        {
            var declarationType = declaration.DeclarationType.ToLocalizedString();
            var declarationName = declaration.IdentifierName;
            return string.Format(
                Resources.Inspections.InspectionResults.IdentifierNameInspection,
                declarationType,
                declarationName);
        }

        protected override ICollection<string> DisabledQuickFixes(Declaration declaration)
        {
            return declaration.DeclarationType.HasFlag(DeclarationType.Module)
                   || declaration.DeclarationType.HasFlag(DeclarationType.Project)
                   || declaration.DeclarationType.HasFlag(DeclarationType.Control)
                   ? new List<string> {nameof(QuickFixes.Concrete.IgnoreOnceQuickFix)}
                   : new List<string>();
        }
    }
}
