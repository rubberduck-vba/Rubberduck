using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.JunkDrawer.Extensions;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.Resources;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Identifies obsolete 16-bit integer variables.
    /// </summary>
    /// <why>
    /// Modern processors are optimized for processing 32-bit integers; internally, a 16-bit integer is still stored as a 32-bit value.
    /// Unless code is interacting with APIs that require a 16-bit integer, a Long (32-bit integer) should be used instead.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim rowCount As Integer
    ///     rowCount = Sheet1.Rows.Count ' overflow: maximum 16-bit signed integer value is only 32,767 (2^15-1)!
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim rowCount As Long
    ///     rowCount = Sheet1.Rows.Count ' all good: maximum 32-bit signed integer value is 2,147,483,647 (2^31-1)!
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class IntegerDataTypeInspection : InspectionBase
    {
        public IntegerDataTypeInspection(RubberduckParserState state) : base(state)
        {
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var interfaceImplementationMembers = State.DeclarationFinder.FindAllInterfaceImplementingMembers().ToHashSet();

            var excludeParameterMembers = State.DeclarationFinder.FindEventHandlers().ToHashSet();
            excludeParameterMembers.UnionWith(interfaceImplementationMembers);

            var result = UserDeclarations
                .Where(declaration =>
                    declaration.AsTypeName == Tokens.Integer &&
                    !interfaceImplementationMembers.Contains(declaration) &&
                    declaration.DeclarationType != DeclarationType.LibraryFunction &&
                    (declaration.DeclarationType != DeclarationType.Parameter || IncludeParameterDeclaration(declaration, excludeParameterMembers)))
                .Select(issue =>
                    new DeclarationInspectionResult(this,
                        string.Format(Resources.Inspections.InspectionResults.IntegerDataTypeInspection,
                            RubberduckUI.ResourceManager.GetString($"DeclarationType_{issue.DeclarationType}", CultureInfo.CurrentUICulture), issue.IdentifierName),
                        issue));

            return result;
        }

        private static bool IncludeParameterDeclaration(Declaration parameterDeclaration, ICollection<Declaration> parentDeclarationsToExclude)
        {
            var parentDeclaration = parameterDeclaration.ParentDeclaration;

            return parentDeclaration.DeclarationType != DeclarationType.LibraryFunction &&
                   parentDeclaration.DeclarationType != DeclarationType.LibraryProcedure &&
                   !parentDeclarationsToExclude.Contains(parentDeclaration);
        }
    }
}
