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

        protected override bool IsUnsuitableArgument(ArgumentReference reference, DeclarationFinder finder)
        {
            var parameter = GetParameterForReference(reference, finder);

            return parameter != null
                    && (!parameter.IsOptional
                        || !parameter.AsTypeName.Equals(Tokens.Variant)
                        || !string.IsNullOrEmpty(parameter.DefaultValue)
                        || parameter.IsArray);
        }

        protected override IInspectionResult InspectionResult(IdentifierReference reference, IDeclarationFinderProvider declarationFinderProvider)
        {
            //This is not ideal, put passing along the declaration finder used in the test itself or the parameter requires reimplementing half of the base class. 
            var argumentReference = reference as ArgumentReference;
            var finder = declarationFinderProvider.DeclarationFinder; 
            var parameter = GetParameterForReference(argumentReference, finder);

            return new IdentifierReferenceInspectionResult(
                this,
                ResultDescription(reference),
                declarationFinderProvider,
                reference,
                parameter);
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            return InspectionResults.IsMissingOnInappropriateArgumentInspection;
        }
    }
}
