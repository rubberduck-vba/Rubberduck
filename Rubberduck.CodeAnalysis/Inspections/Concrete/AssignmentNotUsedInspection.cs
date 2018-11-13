using System.Collections.Generic;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Inspections.CodePathAnalysis;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Inspections.CodePathAnalysis.Extensions;
using System.Linq;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections.Concrete
{
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
                var tree = _walker.GenerateTree(variable.ParentScopeDeclaration.Context, variable);

                nodes.AddRange(tree.GetIdentifierReferences());
            }

            return nodes
                .Select(issue => new IdentifierReferenceInspectionResult(this, Description, State, issue))
                .ToList();
        }
    }
}
