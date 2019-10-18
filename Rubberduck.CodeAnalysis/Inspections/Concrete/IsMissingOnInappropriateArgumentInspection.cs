using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Inspections.Abstract;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Identifies uses of 'IsMissing' involving non-variant, non-optional, or array parameters.
    /// </summary>
    /// <why>
    /// 'IsMissing' only returns True when an optional Variant parameter was not supplied as an argument.
    /// This inspection flags uses that attempt to use 'IsMissing' for other purposes, resulting in conditions that are always False.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Long = 0)
    ///     If IsMissing(foo) Then Exit Sub ' condition is always false
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething(Optional ByVal foo As Variant = 0)
    ///     If IsMissing(foo) Then Exit Sub
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    public class IsMissingOnInappropriateArgumentInspection : IsMissingInspectionBase
    {
        public IsMissingOnInappropriateArgumentInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var results = new List<IInspectionResult>();

            var prefilteredReferences = IsMissingDeclarations.SelectMany(decl => decl.References
                .Where(candidate => !candidate.IsIgnoringInspectionResultFor(AnnotationName)));

            foreach (var reference in prefilteredReferences)
            {
                var parameter = GetParameterForReference(reference);

                if (parameter == null || 
                    parameter.IsOptional 
                    && parameter.AsTypeName.Equals(Tokens.Variant) 
                    && string.IsNullOrEmpty(parameter.DefaultValue) 
                    && !parameter.IsArray)
                {
                    continue;                   
                }

                results.Add(new IdentifierReferenceInspectionResult(this, InspectionResults.IsMissingOnInappropriateArgumentInspection, State, reference, parameter));
            }

            return results;
        }
    }
}
