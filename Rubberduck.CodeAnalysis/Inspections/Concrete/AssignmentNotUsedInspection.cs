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
using Rubberduck.CodeAnalysis.CodePathAnalysis.Execution.ExtendedNodeVisitor;
using Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Warns about a variable that is assigned, and then re-assigned before the first assignment is read.
    /// </summary>
    /// <why>
    /// The first assignment is likely redundant, since it is being overwritten by the second.
    /// </why>
    /// <example hasresult="true">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim foo As Long
    ///     foo = 12 ' assignment is redundant
    ///     foo = 34 
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasresult="false">
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
        IExtendedNodeVisitorFactory _visitorFactory;

        public AssignmentNotUsedInspection(RubberduckParserState state, IExtendedNodeVisitorFactory factory)
            : base(state) 
        {
            _visitorFactory = factory;
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var procedures = State.DeclarationFinder
                    .UserDeclarations(DeclarationType.Member)
                    .Cast<ModuleBodyElementDeclaration>();

            var variables = State.DeclarationFinder
                    .UserDeclarations(DeclarationType.Variable)
                    .Where(d => !d.IsArray)
                .Cast<VariableDeclaration>();

            var nodes = new List<IdentifierReference>();
            foreach (var procedure in procedures)
            {                
                var visitor = _visitorFactory.Create(procedure, State);
                var paths = visitor.GetAllCodePaths();

                foreach (var variable in variables)
                {
                    if (variable.Accessibility == Accessibility.Static
                        || variable.IsIgnoringInspectionResultFor(AnnotationName))
                    {
                        // ignore module-level and static variables... for now
                        // ..also bail out if inspection is ignored at the declaration level.
                        continue;
                    }

                    var assignments = from path in paths
                                      from assignment in path.UnreferencedAssignments
                                      where assignment.IsReachable
                                         && assignment.Target.Declaration.Equals(variable)
                                         && !IsSetAssignmentToNothing(assignment)
                                         && !assignment.Target.IsIgnoringInspectionResultFor(AnnotationName)
                                      select assignment.Target;
                    nodes.AddRange(assignments);
                }
            }

            var results = nodes
                .Select(issue => new IdentifierReferenceInspectionResult(this, Description, State, issue))
                .ToList();
            return results;

            bool IsSetAssignmentToNothing(IAssignmentNode node) =>
                node.Target.Context.Parent is VBAParser.SetStmtContext setStmt &&
                setStmt.expression().GetText().Equals(Tokens.Nothing);
        }
    }
}
