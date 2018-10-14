using System.Collections.Generic;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Inspections.CodePathAnalysis;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Inspections.CodePathAnalysis.Nodes;
using Rubberduck.Inspections.CodePathAnalysis.Extensions;
using System.Linq;
using Rubberduck.Inspections.Results;

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
            var variables = State.DeclarationFinder.UserDeclarations(DeclarationType.Variable);

            var nodes = new List<IdentifierReference>();
            foreach (var variable in variables)
            {
                var tree = _walker.GenerateTree(variable.ParentScopeDeclaration.Context, variable);

                nodes.AddRange(GetIdentifierReferences(tree, variable));
            }

            return nodes
                .Select(issue => new IdentifierReferenceInspectionResult(this, Description, State, issue))
                .ToList();
        }

        private List<IdentifierReference> GetIdentifierReferences(INode node, Declaration declaration)
        {
            var nodes = new List<IdentifierReference>();

            var blockNodes = node.GetNodes(new[] { typeof(BlockNode) });
            foreach (var block in blockNodes)
            {
                INode lastNode = default;
                foreach (var flattenedNode in block.GetFlattenedNodes(new[] { typeof(GenericNode), typeof(BlockNode) }))
                {
                    if (flattenedNode is AssignmentNode &&
                        lastNode is AssignmentNode)
                    {
                        nodes.Add(lastNode.Reference);
                    }

                    lastNode = flattenedNode;
                }

                if (lastNode is AssignmentNode &&
                    block.Children[0].GetFirstNode(new[] { typeof(GenericNode) }) is DeclarationNode)
                {
                    nodes.Add(lastNode.Reference);
                }
            }

            return nodes;
        }
    }
}
