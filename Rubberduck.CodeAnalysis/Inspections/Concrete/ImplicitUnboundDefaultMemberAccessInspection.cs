using System.Collections.Generic;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Parsing.Inspections;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Identifies the use of indexed default member accesses for which the default member cannot be determined at compile time.
    /// </summary>
    /// <why>
    /// Default member accesses hide away the actually called member. This is especially misleading if there is no indication in the expression that such a call is made
    /// and if the default member cannot be determined from the declared type of the object. As a consequence, errors in which a member was forgotten to be called can go unnoticed
    /// and should there not be a suitable default member at runtime, an error 438 'Object doesn't support this property or method' will be raised.
    /// </why>
    /// <example hasresult="true">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal arg As Object)
    ///     Dim bar As Variant
    ///     bar = arg
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasresult="true">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal arg As Object)
    ///     Dim bar As Variant
    ///     arg = bar
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasresult="false">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal arg As Object)
    ///     Dim bar As Variant
    ///     bar = arg.SomeValueReturningMember
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasresult="false">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal arg As Object)
    ///     Dim bar As Variant
    ///     arg.SomePropertyLet = bar
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class ImplicitUnboundDefaultMemberAccessInspection : IdentifierReferenceInspectionBase
    {
        public ImplicitUnboundDefaultMemberAccessInspection(RubberduckParserState state)
            : base(state)
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        protected override IEnumerable<IdentifierReference> ReferencesInModule(QualifiedModuleName module)
        {
            return DeclarationFinderProvider.DeclarationFinder.UnboundDefaultMemberAccesses(module);
        }

        protected override bool IsResultReference(IdentifierReference reference)
        {
            return reference.IsNonIndexedDefaultMemberAccess
                   && !reference.IsProcedureCoercion
                   && !reference.IsInnerRecursiveDefaultMemberAccess
                   && !reference.IsIgnoringInspectionResultFor(AnnotationName);
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            var expression = reference.IdentifierName;
            return string.Format(InspectionResults.ImplicitUnboundDefaultMemberAccessInspection, expression);
        }
    }
}