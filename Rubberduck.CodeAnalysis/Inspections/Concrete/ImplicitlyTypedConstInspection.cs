using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Warns about constants that don't have an explicitly defined type.
    /// </summary>
    /// <why>
    /// All constants have a declared type, whether a type is specified or not. The implicit type is determined by the compiler based on the value, which is not always the expected type.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Const myInteger = 12345
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Const myInteger As Integer = 12345
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Const myInteger% = 12345
    /// ]]>
    /// </example>
    public sealed class ImplicitlyTypedConstInspection : ImplicitTypeInspectionBase
    {
        public ImplicitlyTypedConstInspection(RubberduckParserState state)
            : base(state, DeclarationType.Constant)
        {}

        protected override string ResultDescription(Declaration declaration)
        {
           return string.Format(InspectionResults.ImplicitlyTypedConstInspection, declaration.IdentifierName);
        }
    }
}
