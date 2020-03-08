using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Identifies uses of 'IsMissing' involving a non-parameter argument.
    /// </summary>
    /// <why>
    /// 'IsMissing' only returns True when an optional Variant parameter was not supplied as an argument.
    /// This inspection flags uses that attempt to use 'IsMissing' for other purposes, resulting in conditions that are always False.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim foo As Variant
    ///     If IsMissing(foo) Then Exit Sub ' condition is always false
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(Optional ByVal foo As Variant = 0)
    ///     If IsMissing(foo) Then Exit Sub
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal class IsMissingWithNonArgumentParameterInspection : IsMissingInspectionBase
    {
        public IsMissingWithNonArgumentParameterInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        protected override (bool isResult, ParameterDeclaration properties) IsUnsuitableArgumentWithAdditionalProperties(ArgumentReference reference, DeclarationFinder finder)
        {
            var parameter = ParameterForReference(reference, finder);

            return (parameter == null, null);
        }

        protected override string ResultDescription(IdentifierReference reference, ParameterDeclaration properties)
        {
            return InspectionResults.IsMissingWithNonArgumentParameterInspection;
        }
    }
}
