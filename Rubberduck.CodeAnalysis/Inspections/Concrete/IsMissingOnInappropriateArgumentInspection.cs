using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Inspections.Abstract;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
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

        protected override (bool isResult, object properties) IsUnsuitableArgumentWithAdditionalProperties(ArgumentReference reference, DeclarationFinder finder)
        {
            var parameter = ParameterForReference(reference, finder);

            var isResult = parameter != null
                           && (!parameter.IsOptional
                               || !parameter.AsTypeName.Equals(Tokens.Variant)
                               || !string.IsNullOrEmpty(parameter.DefaultValue)
                               || parameter.IsArray);
            return (isResult, parameter);
        }

        protected override bool IsUnsuitableArgument(ArgumentReference reference, DeclarationFinder finder)
        {
            //No need to implement this, since we override IsUnsuitableArgumentWithAdditionalProperties.
            throw new System.NotImplementedException();
        }

        protected override string ResultDescription(IdentifierReference reference, dynamic properties = null)
        {
            return InspectionResults.IsMissingOnInappropriateArgumentInspection;
        }
    }
}
