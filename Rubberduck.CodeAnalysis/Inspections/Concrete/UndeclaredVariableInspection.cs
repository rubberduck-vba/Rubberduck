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
    /// Warns about implicit local variables that are used but never declared.
    /// </summary>
    /// <why>
    /// If this code compiles, then Option Explicit is omitted and compile-time validation is easily forfeited, even accidentally (e.g. typos). 
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     foo = 42 ' foo is not declared
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim foo As Long
    ///     foo = 42
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class UndeclaredVariableInspection : InspectionBase
    {
        public UndeclaredVariableInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return State.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(item => item.IsUndeclared)
                .Select(item => new DeclarationInspectionResult(this, string.Format(InspectionResults.UndeclaredVariableInspection, item.IdentifierName), item));
        }
    }
}