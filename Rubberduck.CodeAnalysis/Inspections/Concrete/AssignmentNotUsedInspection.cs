using System.Collections.Generic;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Inspections.CodePathAnalysis;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Inspections.CodePathAnalysis.Extensions;
using System.Linq;
using Rubberduck.Inspections.CodePathAnalysis.Nodes;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Warns about a variable that is assigned, and then re-assigned before the first assignment is read.
    /// </summary>
    /// <why>
    /// The first assignment is likely redundant, since it is being overwritten by the second.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim foo As Long
    ///     foo = 12 ' assignment is redundant
    ///     foo = 34 
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Long)
    ///     Dim bar As Long
    ///     bar = 12
    ///     bar = bar + foo ' variable is re-assigned, but the prior assigned value is read at least once first.
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class AssignmentNotUsedInspection : InspectionBase
    {
        private readonly ProcedureTreeVisitor _procedureTreeVisitor;

        public AssignmentNotUsedInspection(RubberduckParserState state, ProcedureTreeVisitor procedureTreeVisitor)
            : base(state) {
            _procedureTreeVisitor = procedureTreeVisitor;
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var variables = State.DeclarationFinder
                    .UserDeclarations(DeclarationType.Variable)
                    .Where(d => !d.IsArray)
                .Cast<VariableDeclaration>();

            var nodes = new List<IdentifierReference>();
            foreach (var procedure in variables.GroupBy(v => v.QualifiedName))
            {
                var scope = procedure.Key;
                var state = new ProcedureTreeVisitorState(State, scope);
                var tree = _procedureTreeVisitor.GenerateTree(scope, state);
                // todo: actually walk the tree
                foreach (var variable in procedure)
                {
                    var parentScopeDeclaration = variable.ParentScopeDeclaration;
                    if (variable.Accessibility == Accessibility.Static)
                    {
                        // ignore module-level and static variables... for now
                        continue;
                    }

                    var assignments = state.Assignments(variable);

                    var references = assignments
                        .Where(node => !node.Usages.Any() &&
                                       !IsSetAssignmentToNothing(node))
                        .Select(node => node.Reference);

                    nodes.AddRange(references);
                }
            }

            var results = nodes
                .Where(issue => !issue.IsIgnoringInspectionResultFor(AnnotationName)
                            // Ignoring the Declaration disqualifies all assignments
                            && !issue.Declaration.IsIgnoringInspectionResultFor(AnnotationName))
                .Select(issue => new IdentifierReferenceInspectionResult(this, Description, State, issue))
                .ToList();
            return results;

            bool IsSetAssignmentToNothing(AssignmentNode node) =>
                node.Reference.Context.Parent is VBAParser.SetStmtContext setStmt &&
                setStmt.expression().GetText().Equals(Tokens.Nothing);
        }
    }
}
