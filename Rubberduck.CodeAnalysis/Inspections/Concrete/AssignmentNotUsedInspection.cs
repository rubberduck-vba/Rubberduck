using System.Collections.Generic;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Inspections.CodePathAnalysis;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Inspections.CodePathAnalysis.Nodes;
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

                nodes.AddRange(GetReferences(tree, variable));
            }

            return nodes
                .Select(issue => new IdentifierReferenceInspectionResult(this, Description, State, issue))
                .ToList();
        }

        private List<IdentifierReference> GetReferences(INode node, Declaration declaration)
        {
            var nodes = new List<IdentifierReference>();

            var blockNodes = GetBlocks(node);
            foreach (var block in blockNodes)
            {
                INode lastNode = default;
                foreach (var flattenedNode in GetFlattenedNodes(block))
                {
                    if (flattenedNode is AssignmentNode &&
                        lastNode is AssignmentNode)
                    {
                        nodes.Add(lastNode.Reference);
                    }

                    lastNode = flattenedNode;
                }

                if (lastNode is AssignmentNode &&
                    GetFirstNonGenericNode(block.Children[0]) is DeclarationNode)
                {
                    nodes.Add(lastNode.Reference);
                }
            }

            return nodes;
        }

        private IEnumerable<INode> GetFlattenedNodes(INode node)
        {
            foreach (var child in node.Children)
            {
                switch (child)
                {
                    case GenericNode _:
                    case BlockNode _:
                        foreach (var nextChild in GetFlattenedNodes(child))
                        {
                            yield return nextChild;
                        }
                        break;
                    default:
                        yield return child;
                        break;
                }
            }
        }

        private IEnumerable<INode> GetBlocks(INode node)
        {
            if (node is BlockNode)
            {
                yield return node;
            }

            foreach (var child in node.Children)
            {
                foreach (var block in GetBlocks(child))
                {
                    yield return block;
                }
            }
        }

        private INode GetFirstNonGenericNode(INode parent)
        {
            if (!(parent is GenericNode))
            {
                return parent;
            }

            return GetFirstNonGenericNode(parent.Children[0]);
        }
    }
}
