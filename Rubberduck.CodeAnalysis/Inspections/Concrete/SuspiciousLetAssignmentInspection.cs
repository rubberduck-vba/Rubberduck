using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Identifies assignments without Set for which both sides are objects.
    /// </summary>
    /// <why>
    /// Whenever both sides of an assignment without Set are objects, there is an assignment from the default member of the RHS to the one on the LHS.
    /// Although this might be intentional, in many situations it will just mask an erroneously forgotten Set. 
    /// </why>
    /// <example hasResult="true">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal rng As Excel.Range, ByVal arg As ADODB Field)
    ///     rng = arg
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResult="false">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal rng As Excel.Range, ByVal arg As ADODB Field)
    ///     rng.Value = arg.Value
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResult="false">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal rng As Excel.Range, ByVal arg As ADODB Field)
    ///     Let rng = arg
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class SuspiciousLetAssignmentInspection : InspectionBase
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public SuspiciousLetAssignmentInspection(RubberduckParserState state)
            : base(state)
        {
            _declarationFinderProvider = state;
            Severity = CodeInspectionSeverity.Warning;
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var results = new List<IInspectionResult>();
            foreach (var moduleDeclaration in State.DeclarationFinder.UserDeclarations(DeclarationType.Module))
            {
                if (moduleDeclaration == null || moduleDeclaration.IsIgnoringInspectionResultFor(AnnotationName))
                {
                    continue;
                }

                var module = moduleDeclaration.QualifiedModuleName;
                results.AddRange(DoGetInspectionResults(module));
            }

            return results;
        }

        private IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module)
        {
            var finder = _declarationFinderProvider.DeclarationFinder;
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
                    var result = InspectionResult(assignment, rhsDefaultMemberAccess, isUnbound, _declarationFinderProvider);
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

        private IInspectionResult InspectionResult(IdentifierReference lhsReference, IdentifierReference rhsReference, bool isUnbound, IDeclarationFinderProvider declarationFinderProvider)
        {
            var result = new IdentifierReferenceInspectionResult(
                this,
                ResultDescription(lhsReference, rhsReference),
                declarationFinderProvider,
                lhsReference);
            result.Properties.RhSReference = rhsReference;
            if (isUnbound)
            {
                result.Properties.DisableFixes = "ExpandDefaultMemberQuickFix";
            }

            return result;
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
                    var result = InspectionResult(assignment, rhsDefaultMemberAccess, true, _declarationFinderProvider);
                    results.Add(result);
                }
            }

            return results;
        }
    }
}