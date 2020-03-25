using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Identifies redundant ByRef modifiers.
    /// </summary>
    /// <why>
    /// Out of convention or preference, explicit ByRef modifiers could be considered redundant since they are the implicit default. 
    /// This inspection can ensure the consistency of the convention.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething(ByRef foo As Long)
    ///     foo = foo + 17
    ///     Debug.Print foo
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// Public Sub DoSomething(foo As Long)
    ///     foo = foo + 17
    ///     Debug.Print foo
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class RedundantByRefModifierInspection : DeclarationInspectionBase
    {
        public RedundantByRefModifierInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, DeclarationType.Parameter)
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            if (!(declaration is ParameterDeclaration parameter)
                || parameter.IsImplicitByRef
                || !parameter.IsByRef
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
                InspectionResults.RedundantByRefModifierInspection,
                declaration.IdentifierName);
        }
    }
}
