using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Warns about constants that don't have an explicitly defined type.
    /// </summary>
    /// <why>
    /// All constants have a declared type, whether a type is specified or not. The implicit type is determined by the compiler based on the value, which is not always the expected type.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Const myInteger = 12345
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Const myInteger As Integer = 12345
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Const myInteger% = 12345
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class ImplicitlyTypedConstInspection : ImplicitTypeInspectionBase
    {
        public ImplicitlyTypedConstInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, DeclarationType.Constant)
        {}

        protected override string ResultDescription(Declaration declaration)
        {
           return string.Format(InspectionResults.ImplicitlyTypedConstInspection, declaration.IdentifierName);
        }
    }
}
