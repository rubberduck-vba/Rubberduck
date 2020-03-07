using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Warns about 'Function' and 'Property Get' procedures that don't have an explicit return type.
    /// </summary>
    /// <why>
    /// All functions return something, whether a type is specified or not. The implicit default is 'Variant'.
    /// </why>
    /// <example hasResult="true">
    /// <![CDATA[
    /// Public Function GetFoo()
    ///     GetFoo = 42
    /// End Function
    /// ]]>
    /// </example>
    /// <example hasResult="false">
    /// <![CDATA[
    /// Public Function GetFoo() As Long
    ///     GetFoo = 42
    /// End Function
    /// ]]>
    /// </example>
    internal sealed class ImplicitVariantReturnTypeInspection : ImplicitTypeInspectionBase
    {
        public ImplicitVariantReturnTypeInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, DeclarationType.Function)
        {}

        protected override string ResultDescription(Declaration declaration)
        {
            return string.Format(InspectionResults.ImplicitVariantReturnTypeInspection, declaration.IdentifierName);
        }
    }
}
