using System.Collections.Generic;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Extensions;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
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
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal arg As Object)
    ///     Dim bar As Variant
    ///     bar = arg
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasresult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal arg As Object)
    ///     Dim bar As Variant
    ///     arg = bar
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasresult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal arg As Object)
    ///     Dim bar As Variant
    ///     bar = arg.SomeValueReturningMember
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasresult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal arg As Object)
    ///     Dim bar As Variant
    ///     arg.SomePropertyLet = bar
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class ImplicitUnboundDefaultMemberAccessInspection : IdentifierReferenceInspectionBase
    {
        public ImplicitUnboundDefaultMemberAccessInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        protected override IEnumerable<IdentifierReference> ReferencesInModule(QualifiedModuleName module, DeclarationFinder finder)
        {
            return finder.UnboundDefaultMemberAccesses(module);
        }

        protected override bool IsResultReference(IdentifierReference reference, DeclarationFinder finder)
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