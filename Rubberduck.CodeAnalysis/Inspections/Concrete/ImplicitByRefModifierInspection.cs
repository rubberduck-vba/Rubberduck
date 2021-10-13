using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using System.Linq;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Highlights implicit ByRef modifiers in user code.
    /// </summary>
    /// <why>
    /// VBA parameters are implicitly ByRef, which differs from modern VB (VB.NET) and most other programming languages which are implicitly ByVal.
    /// So, explicitly identifying VBA parameter mechanisms (the ByRef and ByVal modifiers) can help surface potentially unexpected language results.
    /// The inspection does not flag an implicit parameter mechanism for the last parameter of Property mutators (Let or Set).
    /// VBA applies a ByVal parameter mechanism to the last parameter in the absence (or presence!) of a modifier. 
    /// Exception: UserDefinedType parameters must always be passed as ByRef.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(foo As Long)
    ///     foo = 42
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByRef foo As Long)
    ///     foo = 42
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Private theLength As Long
    /// Public Property Let Length(newLength As Long)
    ///     theLength = newLength
    /// End Sub
    /// ]]>
    /// </module>
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
                || parameter.IsParamArray
                //Exclude parameters of Declare statements
                || !(parameter.ParentDeclaration is ModuleBodyElementDeclaration enclosingMethod))
            {
                return false;
            }

            return !IsPropertyMutatorRHSParameter(enclosingMethod, parameter)
                    && !enclosingMethod.IsInterfaceImplementation
                    && !finder.FindEventHandlers().Contains(enclosingMethod);
        }

        private static bool IsPropertyMutatorRHSParameter(ModuleBodyElementDeclaration enclosingMethod, ParameterDeclaration implicitByRefParameter)
        {
            return (enclosingMethod.DeclarationType.HasFlag(DeclarationType.PropertyLet)
                    || enclosingMethod.DeclarationType.HasFlag(DeclarationType.PropertySet)) 
                && enclosingMethod.Parameters.Last().Equals(implicitByRefParameter);
        }

        protected override string ResultDescription(Declaration declaration)
        {
            return string.Format(
                InspectionResults.ImplicitByRefModifierInspection,
                declaration.IdentifierName);
        }
    }
}
