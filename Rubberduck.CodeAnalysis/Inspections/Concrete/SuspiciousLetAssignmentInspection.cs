using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Extensions;
using Rubberduck.CodeAnalysis.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Identifies assignments without Set for which both sides are objects.
    /// </summary>
    /// <why>
    /// Whenever both sides of an assignment without Set are objects, there is an assignment from the default member of the RHS to the one on the LHS.
    /// Although this might be intentional, in many situations it will just mask an erroneously forgotten Set. 
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal rng As Excel.Range, ByVal arg As ADODB Field)
    ///     rng = arg
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal rng As Excel.Range, ByVal arg As ADODB Field)
    ///     rng.Value = arg.Value
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal rng As Excel.Range, ByVal arg As ADODB Field)
    ///     Let rng = arg
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class SuspiciousLetAssignmentInspection : InspectionBase
    {
        public SuspiciousLetAssignmentInspection(IDeclarationFinderProvider declarationFinderProvider)
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
            return BoundLhsInspectionResults(module, finder)
                .Concat(UnboundLhsInspectionResults(module, finder));
        }

        private IEnumerable<IInspectionResult> BoundLhsInspectionResults(QualifiedModuleName module, DeclarationFinder finder)
        {
            var implicitDefaultMemberAssignments = finder
                .IdentifierReferences(module)
                .Where(IsImplicitDefaultMemberAssignment)
                .ToList();

            var results = new List<IInspectionResult>();
            foreach (var assignment in implicitDefaultMemberAssignments)
            {
                var (rhsDefaultMemberAccess, isUnbound) = RhsImplicitDefaultMemberAccess(assignment, finder);

                if (rhsDefaultMemberAccess != null)
                {
                    var result = InspectionResult(assignment, rhsDefaultMemberAccess, isUnbound, finder);
                    results.Add(result);
                }
            }

            return results;
        }

        private bool IsImplicitDefaultMemberAssignment(IdentifierReference reference)
        {
            return reference.IsNonIndexedDefaultMemberAccess
                   && reference.IsAssignment
                   && !reference.IsSetAssignment
                   && !reference.HasExplicitLetStatement
                   && !reference.IsIgnoringInspectionResultFor(AnnotationName);
        }

        private (IdentifierReference identifierReference, bool isUnbound) RhsImplicitDefaultMemberAccess(IdentifierReference assignment, DeclarationFinder finder)
        {
            if (!(assignment.Context.Parent is VBAParser.LetStmtContext letStatement))
            {
                return (null, false);
            }

            var rhsSelection = new QualifiedSelection(assignment.QualifiedModuleName, letStatement.expression().GetSelection());

            var boundRhsDefaultMemberAccess = finder.IdentifierReferences(rhsSelection)
                .FirstOrDefault(reference => reference.IsNonIndexedDefaultMemberAccess
                                             && !reference.IsInnerRecursiveDefaultMemberAccess);
            if (boundRhsDefaultMemberAccess != null)
            {
                return (boundRhsDefaultMemberAccess, false);
            }

            var unboundRhsDefaultMemberAccess = finder.UnboundDefaultMemberAccesses(rhsSelection.QualifiedName)
                .FirstOrDefault(reference => reference.IsNonIndexedDefaultMemberAccess
                                             && !reference.IsInnerRecursiveDefaultMemberAccess
                                             && reference.Selection.Equals(rhsSelection.Selection));
            return (unboundRhsDefaultMemberAccess, true);
        }

        private IInspectionResult InspectionResult(IdentifierReference lhsReference, IdentifierReference rhsReference, bool isUnbound, DeclarationFinder finder)
        {
            var disabledQuickFixes = isUnbound
                ? new List<string> {"ExpandDefaultMemberQuickFix"}
                : new List<string>();
            return new IdentifierReferenceInspectionResult<IdentifierReference>(
                this,
                ResultDescription(lhsReference, rhsReference),
                finder,
                lhsReference,
                rhsReference,
                disabledQuickFixes);
        }

        private string ResultDescription(IdentifierReference lhsReference, IdentifierReference rhsReference)
        {
            var lhsExpression = lhsReference.IdentifierName;
            var rhsExpression = rhsReference.IdentifierName;
            return string.Format(InspectionResults.SuspiciousLetAssignmentInspection, lhsExpression, rhsExpression);
        }

        private IEnumerable<IInspectionResult> UnboundLhsInspectionResults(QualifiedModuleName module, DeclarationFinder finder)
        {
            var implicitDefaultMemberAssignments = finder
                .UnboundDefaultMemberAccesses(module)
                .Where(IsImplicitDefaultMemberAssignment);

            var results = new List<IInspectionResult>();
            foreach (var assignment in implicitDefaultMemberAssignments)
            {
                var (rhsDefaultMemberAccess, isUnbound) = RhsImplicitDefaultMemberAccess(assignment, finder);

                if (rhsDefaultMemberAccess != null)
                {
                    var result = InspectionResult(assignment, rhsDefaultMemberAccess, true, finder);
                    results.Add(result);
                }
            }

            return results;
        }
    }
}