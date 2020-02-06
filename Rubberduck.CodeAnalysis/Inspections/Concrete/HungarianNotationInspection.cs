using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Settings;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.SettingsProvider;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Flags identifiers that use [Systems] Hungarian Notation prefixes.
    /// </summary>
    /// <why>
    /// Systems Hungarian (encoding data types in variable names) stemmed from a misunderstanding of what its inventor meant
    /// when they described that prefixes identified the "kind" of variable in a naming scheme dubbed Apps Hungarian.
    /// Modern naming conventions in all programming languages heavily discourage the use of Systems Hungarian prefixes. 
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim bFoo As Boolean, blnFoo As Boolean
    ///     Dim intBar As Long ' which is correct? the int or the Long?
    /// End Sub
    ///
    /// Private Function fnlngGetFoo() As Long
    ///     fnlngGetFoo = 42
    /// End Function
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim foo As Boolean, isFoo As Boolean
    ///     Dim bar As long
    /// End Sub
    /// 
    /// Private Function GetFoo() As Long
    ///     GetFoo = 42
    /// End Function
    /// ]]>
    /// </example>
    public sealed class HungarianNotationInspection : InspectionBase
    {
        private static readonly List<DeclarationType> TargetDeclarationTypes = new List<DeclarationType>
        {
            DeclarationType.Parameter,
            DeclarationType.Constant,
            DeclarationType.Control,
            DeclarationType.ClassModule,
            DeclarationType.Document,
            DeclarationType.Member,
            DeclarationType.Module,
            DeclarationType.ProceduralModule,
            DeclarationType.UserForm,
            DeclarationType.UserDefinedType,
            DeclarationType.UserDefinedTypeMember,
            DeclarationType.Variable
        };

        private static readonly List<DeclarationType> IgnoredProcedureTypes = new List<DeclarationType>
        {
            DeclarationType.LibraryFunction,
            DeclarationType.LibraryProcedure
        };

        private readonly IConfigurationService<CodeInspectionSettings> _settings;

        public HungarianNotationInspection(RubberduckParserState state, IConfigurationService<CodeInspectionSettings> settings)
            : base(state)
        {
            _settings = settings;
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var finder = DeclarationFinderProvider.DeclarationFinder;
            var settings = _settings.Read();
            var whitelistedNames = settings.WhitelistedIdentifiers
                .Select(s => s.Identifier)
                .ToList();

            var results = new List<IInspectionResult>();
            foreach (var moduleDeclaration in State.DeclarationFinder.UserDeclarations(DeclarationType.Module))
            {
                if (moduleDeclaration == null)
                {
                    continue;
                }

                var module = moduleDeclaration.QualifiedModuleName;
                results.AddRange(DoGetInspectionResults(module, finder, whitelistedNames));
            }

            return results;
        }

        private IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module)
        {
            var finder = DeclarationFinderProvider.DeclarationFinder;
            var settings = _settings.Read();
            var whitelistedNames = settings.WhitelistedIdentifiers
                .Select(s => s.Identifier)
                .ToList();
            return DoGetInspectionResults(module, finder, whitelistedNames);
        }

        private IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder, List<string> whitelistedNames)
        {
            var objectionableDeclarations = RelevantDeclarationsInModule(module, finder)
                .Where(declaration => IsResultDeclaration(declaration, whitelistedNames));

            return objectionableDeclarations
                .Select(InspectionResult)
                .ToList();
        }

        private static bool IsResultDeclaration(Declaration declaration, ICollection<string> whitelistedNames)
        {
            return !whitelistedNames.Contains(declaration.IdentifierName)
                   && !IgnoredProcedureTypes.Contains(declaration.ParentDeclaration.DeclarationType)
                   && declaration.IdentifierName.TryMatchHungarianNotationCriteria(out _);
        }

        private IInspectionResult InspectionResult(Declaration declaration)
        {
            return new DeclarationInspectionResult(
                this,
                ResultDescription(declaration),
                declaration);
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

        private IEnumerable<Declaration> RelevantDeclarationsInModule(QualifiedModuleName module, DeclarationFinder finder)
        {
            var potentiallyRelevantDeclarations = TargetDeclarationTypes
                    .SelectMany(declarationType => finder.Members(module, declarationType))
                    .Distinct();
            return potentiallyRelevantDeclarations
                .Where(declaration => !IgnoredProcedureTypes.Contains(declaration.DeclarationType));
        }
    }
}
