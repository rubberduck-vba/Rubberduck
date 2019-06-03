using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Inspections.Abstract;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.Inspections.Inspections.Concrete
{
    /// <summary>
    /// Identifies uses of 'IsMissing' involving a non-parameter argument.
    /// </summary>
    /// <why>
    /// 'IsMissing' only returns True when an optional Variant parameter was not supplied as an argument.
    /// This inspection flags uses that attempt to use 'IsMissing' for other purposes, resulting in conditions that are always False.
    /// </why>
    /// <example>
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim foo As Variant
    ///     If IsMissing(foo) Then Exit Sub ' condition is always false
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    /// <example>
    /// <![CDATA[
    /// Public Sub DoSomething(Optional ByVal foo As Variant = 0)
    ///     If IsMissing(foo) Then Exit Sub
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    public class IsMissingWithNonArgumentParameterInspection : IsMissingInspectionBase
    {
        public IsMissingWithNonArgumentParameterInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var results = new List<IInspectionResult>();

            foreach (var reference in IsMissingDeclarations.SelectMany(decl => decl.References
                .Where(candidate => !candidate.IsIgnoringInspectionResultFor(AnnotationName))))
            {
                var parameter = GetParameterForReference(reference);

                if (parameter != null)
                {
                    continue;
                }

                results.Add(new IdentifierReferenceInspectionResult(this, InspectionResults.IsMissingWithNonArgumentParameterInspection, State, reference));
            }

            return results;
        }
    }
}
