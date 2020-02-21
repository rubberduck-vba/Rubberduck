using Rubberduck.Inspections.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Identifies uses of 'IsMissing' involving a non-parameter argument.
    /// </summary>
    /// <why>
    /// 'IsMissing' only returns True when an optional Variant parameter was not supplied as an argument.
    /// This inspection flags uses that attempt to use 'IsMissing' for other purposes, resulting in conditions that are always False.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim foo As Variant
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
    public class IsMissingWithNonArgumentParameterInspection : IsMissingInspectionBase
    {
        public IsMissingWithNonArgumentParameterInspection(RubberduckParserState state)
            : base(state) { }

        protected override bool IsUnsuitableArgument(ArgumentReference reference, DeclarationFinder finder)
        {
            var parameter = ParameterForReference(reference, finder);

            return parameter == null;
        }

        protected override string ResultDescription(IdentifierReference reference, dynamic properties = null)
        {
            return InspectionResults.IsMissingWithNonArgumentParameterInspection;
        }
    }
}
