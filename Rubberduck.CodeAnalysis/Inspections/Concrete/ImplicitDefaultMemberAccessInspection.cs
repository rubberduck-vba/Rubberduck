using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Extensions;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Identifies the use of non-indexed default member accesses.
    /// </summary>
    /// <why>
    /// Default member accesses hide away the actually called member. This is especially misleading if there is no indication in the expression that such a call is made
    /// and can cause errors in which a member was forgotten to be called to go unnoticed.
    /// </why>
    /// <example hasresult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal arg As ADODB.Field)
    ///     Dim bar As Variant
    ///     bar = arg
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasresult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal arg As ADODB.Connection)
    ///     Dim bar As String
    ///     arg = bar
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasresult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal arg As ADODB.Field)
    ///     Dim bar As Variant
    ///     bar = arg.Value
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasresult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal arg As ADODB.Connection)
    ///     Dim bar As String
    ///     arg.ConnectionString = bar
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class ImplicitDefaultMemberAccessInspection : IdentifierReferenceInspectionBase
    {
        public ImplicitDefaultMemberAccessInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        protected override bool IsResultReference(IdentifierReference reference, DeclarationFinder finder)
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