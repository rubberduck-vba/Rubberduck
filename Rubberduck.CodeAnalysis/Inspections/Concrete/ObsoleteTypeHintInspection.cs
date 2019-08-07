using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Flags declarations where a type hint is used in place of an 'As' clause.
    /// </summary>
    /// <why>
    /// Type hints were made obsolete when declaration syntax introduced the 'As' keyword. Prefer explicit type names over type hint symbols.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim foo$
    ///     foo = "some string"
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim foo As String
    ///     foo = "some string"
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class ObsoleteTypeHintInspection : InspectionBase
    {
        public ObsoleteTypeHintInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var results = UserDeclarations.ToList();

            var declarations = from item in results
                where item.HasTypeHint
                select
                    new DeclarationInspectionResult(this,
                        string.Format(InspectionResults.ObsoleteTypeHintInspection,
                            InspectionsUI.Inspections_Declaration, item.DeclarationType.ToString().ToLower(),
                            item.IdentifierName), item);

            var references = from item in results.SelectMany(d => d.References)
                where item.HasTypeHint()
                select
                    new IdentifierReferenceInspectionResult(this,
                        string.Format(InspectionResults.ObsoleteTypeHintInspection,
                            InspectionsUI.Inspections_Usage, item.Declaration.DeclarationType.ToString().ToLower(),
                            item.IdentifierName),
                        State,
                        item);

            return declarations.Union<IInspectionResult>(references);
        }
    }
}
