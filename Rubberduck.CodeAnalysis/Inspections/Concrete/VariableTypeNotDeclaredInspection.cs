using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Warns about variables declared without an explicit data type.
    /// </summary>
    /// <why>
    /// A variable declared without an explicit data type is implicitly a Variant/Empty until it is assigned.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim value ' implicit Variant
    ///     value = 42
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim value As Long
    ///     value = 42
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class VariableTypeNotDeclaredInspection : InspectionBase
    {
        public VariableTypeNotDeclaredInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var issues = from item in State.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                         .Union(State.DeclarationFinder.UserDeclarations(DeclarationType.Parameter))
                         where (item.DeclarationType != DeclarationType.Parameter || (item.DeclarationType == DeclarationType.Parameter && !item.IsArray))
                         && item.DeclarationType != DeclarationType.Control
                         && !item.IsTypeSpecified
                         && !item.IsUndeclared
                         select new DeclarationInspectionResult(this, string.Format(InspectionResults.ImplicitVariantDeclarationInspection, item.DeclarationType, item.IdentifierName), item);
            return issues;
        }
    }
}
