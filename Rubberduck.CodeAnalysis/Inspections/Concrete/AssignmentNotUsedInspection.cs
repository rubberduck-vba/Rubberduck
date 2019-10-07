using System.Collections.Generic;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Inspections.CodePathAnalysis;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Inspections.CodePathAnalysis.Extensions;
using System.Linq;
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
        private readonly Walker _walker;

        public AssignmentNotUsedInspection(RubberduckParserState state, Walker walker)
            : base(state) {
            _walker = walker;
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var variables = State.DeclarationFinder
                    .UserDeclarations(DeclarationType.Variable)
                    .Where(d => !d.IsArray);

            var nodes = new List<IdentifierReference>();
            foreach (var variable in variables)
            {
                var parentScopeDeclaration = variable.ParentScopeDeclaration;

                if (parentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Module))
                {
                    continue;
                }

                var tree = _walker.GenerateTree(parentScopeDeclaration.Context, variable);

                var references = tree.GetIdentifierReferences();
                // ignore set-assignments to 'Nothing'
                nodes.AddRange(references.Where(r =>
                    !(r.Context.Parent is VBAParser.SetStmtContext setStmtContext &&
                    setStmtContext.expression().GetText().Equals(Tokens.Nothing))));
            }

            return nodes
                // Ignoring the Declaration disqualifies all assignments
                .Where(issue => !issue.Declaration.IsIgnoringInspectionResultFor(AnnotationName))
                .Select(issue => new IdentifierReferenceInspectionResult(this, Description, State, issue))
                .ToList();
        }
    }
}
