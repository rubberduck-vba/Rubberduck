using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Identifies implicit default member calls.
    /// </summary>
    /// <why>
    /// Code should do what it says, and say what it does. Implicit default member calls generally do the opposite of that.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     ActiveSheet.Range("A1") = 42 ' implicit assignment to 'Range.[_default]'.
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     ActiveSheet.Range("A1").Value = 42
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class ImplicitDefaultMemberAssignmentInspection : InspectionBase
    {
        public ImplicitDefaultMemberAssignmentInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var boundDefaultMemberAssignments = State.DeclarationFinder
                .AllIdentifierReferences()
                .Where(IsRelevantReference);

            var boundIssues = boundDefaultMemberAssignments
                .Select(reference => new IdentifierReferenceInspectionResult(
                    this,
                    string.Format(
                        InspectionResults.ImplicitDefaultMemberAssignmentInspection,
                        reference.Context.GetText(),
                        reference.Declaration.IdentifierName,
                        reference.Declaration.QualifiedModuleName.ToString()),
                    State,
                    reference));

            var unboundDefaultMemberAssignments = State.DeclarationFinder
                .AllUnboundDefaultMemberAccesses()
                .Where(IsRelevantReference);

            var unboundIssues = unboundDefaultMemberAssignments
                .Select(reference => new IdentifierReferenceInspectionResult(
                    this,
                    string.Format(
                        InspectionResults.ImplicitDefaultMemberAssignmentInspection_Unbound,
                        reference.Context.GetText()),
                    State,
                    reference));

            return boundIssues.Concat(unboundIssues);
        }

        private bool IsRelevantReference(IdentifierReference reference)
        {
            return reference.IsAssignment
                   && reference.IsNonIndexedDefaultMemberAccess
                   && !reference.IsIgnoringInspectionResultFor(AnnotationName);
        }
    }
}