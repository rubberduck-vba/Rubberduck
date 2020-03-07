using Rubberduck.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Highlights implicit ByRef modifiers in user code.
    /// </summary>
    /// <why>
    /// In modern VB (VB.NET), the implicit modifier is ByVal, as it is in most other programming languages.
    /// Making the ByRef modifiers explicit can help surface potentially unexpected language defaults.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething(foo As Long)
    ///     foo = 42
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething(ByRef foo As Long)
    ///     foo = 42
    /// End Sub
    /// ]]>
    /// </example>
    internal sealed class ImplicitByRefModifierInspection : DeclarationInspectionBase
    {
        public ImplicitByRefModifierInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, DeclarationType.Parameter)
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            if (!(declaration is ParameterDeclaration parameter)
                || !parameter.IsImplicitByRef
                || parameter.IsParamArray)
            {
                return false;
            }

            var parentDeclaration = parameter.ParentDeclaration;

            if (parentDeclaration is ModuleBodyElementDeclaration enclosingMethod)
            {
                return !enclosingMethod.IsInterfaceImplementation
                       && !finder.FindEventHandlers().Contains(enclosingMethod);
            }

            return parentDeclaration.DeclarationType != DeclarationType.LibraryFunction
                   && parentDeclaration.DeclarationType != DeclarationType.LibraryProcedure;
        }

        protected override string ResultDescription(Declaration declaration)
        {
            return string.Format(
                InspectionResults.ImplicitByRefModifierInspection,
                declaration.IdentifierName);
        }
    }
}
