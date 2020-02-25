using Rubberduck.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Warns about 'Function' and 'Property Get' procedures that don't have an explicit return type.
    /// </summary>
    /// <why>
    /// All functions return something, whether a type is specified or not. The implicit default is 'Variant'.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Function GetFoo()
    ///     GetFoo = 42
    /// End Function
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Function GetFoo() As Long
    ///     GetFoo = 42
    /// End Function
    /// ]]>
    /// </example>
    public sealed class ImplicitVariantReturnTypeInspection : ImplicitTypeInspectionBase
    {
        public ImplicitVariantReturnTypeInspection(RubberduckParserState state)
            : base(state, DeclarationType.Function)
        {}

        protected override string ResultDescription(Declaration declaration)
        {
            return string.Format(InspectionResults.ImplicitVariantReturnTypeInspection, declaration.IdentifierName);
        }
    }
}
