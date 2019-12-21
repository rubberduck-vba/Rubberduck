using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Rubberduck.CodeAnalysis.Settings;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources;
using Rubberduck.SettingsProvider;

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
        #region statics
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

        #endregion

        private readonly IConfigurationService<CodeInspectionSettings> _settings;

        public HungarianNotationInspection(RubberduckParserState state, IConfigurationService<CodeInspectionSettings> settings)
            : base(state)
        {
            _settings = settings;
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var settings = _settings.Read();
            var whitelistedNames = settings.WhitelistedIdentifiers.Select(s => s.Identifier).ToList();

            var hungarians = UserDeclarations
                .Where(declaration => !whitelistedNames.Contains(declaration.IdentifierName)
                                      && TargetDeclarationTypes.Contains(declaration.DeclarationType)
                                      && !IgnoredProcedureTypes.Contains(declaration.DeclarationType)
                                      && !IgnoredProcedureTypes.Contains(declaration.ParentDeclaration.DeclarationType)
                                      && declaration.IdentifierName.TryMatchHungarianNotationCriteria(out _))
                .Select(issue => new DeclarationInspectionResult(this,
                                                      string.Format(Resources.Inspections.InspectionResults.IdentifierNameInspection,
                                                                    RubberduckUI.ResourceManager.GetString($"DeclarationType_{issue.DeclarationType}", CultureInfo.CurrentUICulture),
                                                                    issue.IdentifierName),
                                                      issue));

            return hungarians;
        }
    }
}
