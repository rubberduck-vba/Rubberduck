using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Extensions;
using Rubberduck.CodeAnalysis.Settings;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.SettingsProvider;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Flags identifiers that use [Systems] Hungarian Notation prefixes.
    /// </summary>
    /// <why>
    /// Systems Hungarian (encoding data types in variable names) stemmed from a misunderstanding of what its inventor meant
    /// when they described that prefixes identified the "kind" of variable in a naming scheme dubbed Apps Hungarian.
    /// Modern naming conventions in all programming languages heavily discourage the use of Systems Hungarian prefixes. 
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
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
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
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
    /// </module>
    /// </example>
    internal sealed class HungarianNotationInspection : DeclarationInspectionUsingGlobalInformationBase<List<string>>
    {
        private static readonly DeclarationType[] TargetDeclarationTypes = new []
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

        private static readonly DeclarationType[] IgnoredProcedureTypes = new []
        {
            DeclarationType.LibraryFunction,
            DeclarationType.LibraryProcedure
        };

        private readonly IConfigurationService<CodeInspectionSettings> _settings;

        public HungarianNotationInspection(IDeclarationFinderProvider declarationFinderProvider, IConfigurationService<CodeInspectionSettings> settings)
            : base(declarationFinderProvider, TargetDeclarationTypes, IgnoredProcedureTypes)
        {
            _settings = settings;
        }

        protected override List<string> GlobalInformation(DeclarationFinder finder)
        {
            var settings = _settings.Read();
            return settings.WhitelistedIdentifiers
                .Select(s => s.Identifier)
                .ToList();
        }

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder, List<string> whitelistedNames)
        {
            return !whitelistedNames.Contains(declaration.IdentifierName)
                   && !IgnoredProcedureTypes.Contains(declaration.ParentDeclaration.DeclarationType)
                   && declaration.IdentifierName.TryMatchHungarianNotationCriteria(out _);
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
    }
}
