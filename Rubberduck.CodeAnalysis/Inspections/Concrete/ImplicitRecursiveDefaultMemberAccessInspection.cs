using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Extensions;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Identifies the use of indexed default member accesses that require a recursive default member resolution.
    /// </summary>
    /// <why>
    /// Default member accesses hide away the actually called member. This is especially misleading if there is no indication in the expression that such a call is made
    /// and the final default member is not on the interface of the object itself. In particular, this can cause errors in which a member was forgotten to be called to go unnoticed.
    /// </why>
    /// <example hasresult="true">
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// Public Function Foo() As Class2
    /// Attibute Foo.VB_UserMemId = 0
    ///     Set Foo = New Class2
    /// End Function
    /// ]]>
    /// </module>
    /// <module name="Class2" type="Class Module">
    /// <![CDATA[
    /// Public Function Bar() As Long
    /// Attibute Bar.VB_UserMemId = 0
    ///     Foo = 42
    /// End Function
    /// ]]>
    /// </module>
    /// <module name="Module" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal arg As Class1)
    ///     Dim bar As Variant
    ///     bar = arg
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasresult="false">
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// Public Function Foo() As Class2
    /// Attibute Foo.VB_UserMemId = 0
    ///     Set Foo = New Class2
    /// End Function
    /// ]]>
    /// </module>
    /// <module name="Class2" type="Class Module">
    /// <![CDATA[
    /// Public Function Bar() As Long
    /// Attibute Bar.VB_UserMemId = 0
    ///     Foo = 42
    /// End Function
    /// ]]>
    /// </module>
    /// <module name="Module" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal arg As Class1)
    ///     Dim bar As Variant
    ///     bar = arg.Foo().Bar()
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class ImplicitRecursiveDefaultMemberAccessInspection : IdentifierReferenceInspectionBase
    {
        public ImplicitRecursiveDefaultMemberAccessInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        protected override bool IsResultReference(IdentifierReference reference, DeclarationFinder finder)
        {
            return reference.IsNonIndexedDefaultMemberAccess
                   && reference.DefaultMemberRecursionDepth > 1
                   && !reference.IsProcedureCoercion
                   && !reference.IsInnerRecursiveDefaultMemberAccess
                   && !reference.IsIgnoringInspectionResultFor(AnnotationName);
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            var expression = reference.IdentifierName;
            var defaultMember = reference.Declaration.QualifiedName.ToString();
            return string.Format(InspectionResults.ImplicitRecursiveDefaultMemberAccessInspection, expression, defaultMember);
        }
    }
}