using Rubberduck.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Parsing.Inspections;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Identifies the use non-indexed default member accesses.
    /// </summary>
    /// <why>
    /// Default member accesses hide away the actually called member. This is especially misleading if there is no indication in the expression that such a call is made
    /// and can cause errors in which a member was forgotten to be called to go unnoticed.
    /// </why>
    /// <example hasResult="true">
    /// <![CDATA[
    /// Class1:
    ///
    /// Public Function Foo() As Long
    /// Attibute Foo.VB_UserMemId = 0
    ///     Foo = 42
    /// End Function
    ///
    /// Module:
    /// 
    /// Public Sub DoSomething(ByVal arg As Class1)
    ///     Dim bar As Variant
    ///     bar = arg
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResult="true">
    /// <![CDATA[
    /// Class1:
    ///
    /// Public Property Let Foo(RHS As Long)
    /// Attibute Foo.VB_UserMemId = 0
    /// End Function
    ///
    /// Module:
    /// 
    /// Public Sub DoSomething(ByVal arg As Class1)
    ///     Dim bar As Variant
    ///     arg = bar
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResult="false">
    /// <![CDATA[
    /// Class1:
    ///
    /// Public Function Foo() As Long
    /// Attibute Foo.VB_UserMemId = 0
    ///     Foo = 42
    /// End Function
    ///
    /// Module:
    /// 
    /// Public Sub DoSomething(ByVal arg As Class1)
    ///     Dim bar As Variant
    ///     bar = arg.Foo()
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResult="false">
    /// <![CDATA[
    /// Class1:
    ///
    /// Public Property Let Foo(RHS As Long)
    /// Attibute Foo.VB_UserMemId = 0
    /// End Function
    ///
    /// Module:
    /// 
    /// Public Sub DoSomething(ByVal arg As Class1)
    ///     Dim bar As Variant
    ///     arg.Foo = bar
    /// End Sub
    /// ]]>
    public sealed class ImplicitDefaultMemberAccessInspection : IdentifierReferenceInspectionBase
    {
        public ImplicitDefaultMemberAccessInspection(RubberduckParserState state)
            : base(state)
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        protected override bool IsResultReference(IdentifierReference reference)
        {
            return reference.IsNonIndexedDefaultMemberAccess
                   && reference.DefaultMemberRecursionDepth == 1
                   && !reference.IsProcedureCoercion
                   && !reference.IsInnerRecursiveDefaultMemberAccess
                   && !reference.IsIgnoringInspectionResultFor(AnnotationName);
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            var expression = reference.IdentifierName;
            var defaultMember = reference.Declaration.QualifiedName.ToString();
            return string.Format(InspectionResults.ImplicitDefaultMemberAccessInspection, expression, defaultMember);
        }
    }
}