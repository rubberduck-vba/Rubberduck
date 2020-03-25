using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Results;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Identifies places in which an object is used but a procedure is required and a default member exists on the object.
    /// </summary>
    /// <why>
    /// Providing an object where a procedure is required leads to an implicit call to the object's default member.
    /// This behavior is not obvious, and most likely unintended.
    /// </why>
    /// <example hasresult="true">
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// Public Function Foo() As Long
    /// Attibute Foo.VB_UserMemId = 0
    ///     Foo = 42
    /// End Function
    /// ]]>
    /// </module>
    /// <module name="Module" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal arg As Class1)
    ///     arg
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasresult="false">
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// Public Function Foo() As Long
    /// Attibute Foo.VB_UserMemId = 0
    ///     Foo = 42
    /// End Function
    /// ]]>
    /// </module>
    /// <module name="Module" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal arg As Class1)
    ///     arg.Foo
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class ObjectWhereProcedureIsRequiredInspection : InspectionBase
    {
        public ObjectWhereProcedureIsRequiredInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(DeclarationFinder finder)
        {
            return finder.UserDeclarations(DeclarationType.Module)
                .Where(module => module != null)
                .SelectMany(module => DoGetInspectionResults(module.QualifiedModuleName, finder));
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder)
        {
            return BoundInspectionResults(module, finder)
                .Concat(UnboundInspectionResults(module, finder));
        }

        private IEnumerable<IInspectionResult> BoundInspectionResults(QualifiedModuleName module, DeclarationFinder finder)
        {
            var objectionableReferences = finder
                .IdentifierReferences(module)
                .Where(IsResultReference);

            return objectionableReferences
                .Select(reference => BoundInspectionResult(reference, finder))
                .ToList();
        }

        private static bool IsResultReference(IdentifierReference reference)
        {
            return reference.IsProcedureCoercion;
        }

        private IInspectionResult BoundInspectionResult(IdentifierReference reference, DeclarationFinder finder)
        {
            return new IdentifierReferenceInspectionResult(
                this,
                BoundResultDescription(reference),
                finder,
                reference);
        }

        private static string BoundResultDescription(IdentifierReference reference)
        {
            var expression = reference.IdentifierName;
            var defaultMember = reference.Declaration.QualifiedName.ToString();
            return string.Format(InspectionResults.ObjectWhereProcedureIsRequiredInspection, expression, defaultMember);
        }

        private IEnumerable<IInspectionResult> UnboundInspectionResults(QualifiedModuleName module, DeclarationFinder finder)
        {
            var objectionableReferences = finder
                .UnboundDefaultMemberAccesses(module)
                .Where(IsResultReference);

            return objectionableReferences
                .Select(reference => UnboundInspectionResult(reference, finder))
                .ToList();
        }

        private IInspectionResult UnboundInspectionResult(IdentifierReference reference, DeclarationFinder finder)
        {
            var disabledQuickFixes = new List<string>{ "ExpandDefaultMemberQuickFix" };
            return new IdentifierReferenceInspectionResult(
                this,
                UnboundResultDescription(reference),
                finder,
                reference,
                disabledQuickFixes);
        }

        private static string UnboundResultDescription(IdentifierReference reference)
        {
            var expression = reference.IdentifierName;
            return string.Format(InspectionResults.ObjectWhereProcedureIsRequiredInspection_Unbound, expression);
        }
    }
}