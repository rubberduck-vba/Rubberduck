using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using System.Linq;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Flags the value-parameter of a property mutators that are declared with an explict ByRef modifier.
    /// </summary>
    /// <why>
    /// Regardless of the presence or absence of an explicit ByRef or ByVal modifier, the value-parameter
    /// of a property mutator is always treated as though it had an explicit ByVal modifier.
    /// Exception: UserDefinedType and Array parameters are always passed by reference.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Private fizzField As Long
    /// Public Property Get Fizz() As Long
    ///     Fizz = fizzField
    /// End Property
    /// Public Property Let Fizz(ByRef arg As Long)
    ///     fizzField = arg
    /// End Property
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Private fizzField As Long
    /// Public Property Get Fizz() As Long
    ///     Fizz = fizzField
    /// End Property
    /// Public Property Let Fizz(arg As Long)
    ///     fizzField = arg
    /// End Property
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class MisleadingByRefParameterInspection : DeclarationInspectionBase
    {
        public MisleadingByRefParameterInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, DeclarationType.Parameter)
        { }

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            return declaration is ParameterDeclaration parameter
                && !IsAlwaysByRef(declaration)
                && declaration.ParentDeclaration is ModuleBodyElementDeclaration enclosingMethod
                && (enclosingMethod.DeclarationType.HasFlag(DeclarationType.PropertyLet)
                    || enclosingMethod.DeclarationType.HasFlag(DeclarationType.PropertySet))
                && enclosingMethod.Parameters.Last() == parameter
                && parameter.IsByRef && !parameter.IsImplicitByRef;
        }

        private static bool IsAlwaysByRef(Declaration parameter) 
            => parameter.IsArray
                || (parameter.AsTypeDeclaration?.DeclarationType.HasFlag(DeclarationType.UserDefinedType) ?? false);

        protected override string ResultDescription(Declaration declaration)
        {
            return string.Format(
                InspectionResults.MisleadingByRefParameterInspection,
                declaration.IdentifierName, declaration.ParentDeclaration.QualifiedName.MemberName);
        }
    }
}
