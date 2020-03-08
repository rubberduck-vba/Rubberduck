using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Extensions;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Warns about variables declared without an explicit data type.
    /// </summary>
    /// <why>
    /// A variable declared without an explicit data type is implicitly a Variant/Empty until it is assigned.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim value ' implicit Variant
    ///     value = 42
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim value As Long
    ///     value = 42
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class VariableTypeNotDeclaredInspection : ImplicitTypeInspectionBase
    {
        public VariableTypeNotDeclaredInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, new []{DeclarationType.Parameter, DeclarationType.Variable}, new[]{DeclarationType.Control})
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            return base.IsResultDeclaration(declaration, finder)
                   && !declaration.IsUndeclared
                   && (declaration.DeclarationType != DeclarationType.Parameter 
                       || declaration is ParameterDeclaration parameter && !parameter.IsParamArray);
        }

        protected override string ResultDescription(Declaration declaration)
        {
            var declarationType = declaration.DeclarationType.ToLocalizedString();
            var declarationName = declaration.IdentifierName;
            return string.Format(
                InspectionResults.ImplicitVariantDeclarationInspection,
                declarationType,
                declarationName);
        }
    }
}
