using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Warns about variables declared without an explicit data type.
    /// </summary>
    /// <why>
    /// A variable declared without an explicit data type is implicitly a Variant/Empty until it is assigned.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim value ' implicit Variant
    ///     value = 42
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim value As Long
    ///     value = 42
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class VariableTypeNotDeclaredInspection : ImplicitTypeInspectionBase
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
