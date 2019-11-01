using Rubberduck.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Parsing.Inspections;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Identifies the use of non-indexed default member accesses.
    /// </summary>
    /// <why>
    /// Default member accesses hide away the actually called member. This is especially misleading if there is no indication in the expression that such a call is made
    /// and can cause errors in which a member was forgotten to be called to go unnoticed.
    /// </why>
    /// <example hasresult="true">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal arg As ADODB.Field)
    ///     Dim bar As Variant
    ///     bar = arg
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasresult="true">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal arg As ADODB.Connection)
    ///     Dim bar As String
    ///     arg = bar
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasresult="false">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal arg As ADODB.Field)
    ///     Dim bar As Variant
    ///     bar = arg.Value
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasresult="false">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal arg As ADODB.Connection)
    ///     Dim bar As String
    ///     arg.ConnectionString = bar
    /// End Sub
    /// ]]>
    /// </example>
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