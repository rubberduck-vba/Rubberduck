using Rubberduck.Inspections.CodePathAnalysis.Nodes;
using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.CodePathAnalysis.Extensions
{
    public static class NodeExtensions
    {
        public static IEnumerable<INode> GetFlattenedNodes(this INode node, params Type[] excludedTypes)
        {
            foreach (var child in node.Children)
            {
                if (!excludedTypes.Contains(child.GetType()))
                {
                    yield return child;
                }
                else
                {
                    foreach (var nextChild in GetFlattenedNodes(child, excludedTypes))
                    {
                        yield return nextChild;
                    }
                }
            }
        }

        public static IEnumerable<INode> GetNodes(this INode node, params Type[] types)
        {
            if (types.Contains(node.GetType()))
            {
                yield return node;
            }

            foreach (var child in node.Children)
            {
                foreach (var childNode in GetNodes(child, types))
                {
                    yield return childNode;
                }
            }
        }

        /// <summary>
        /// Starting with the specified root, recursively searches for the first node in a node tree that isn't one of the optionally specified excluded types.
        /// </summary>
        /// <param name="node"></param>
        /// <param name="excludedTypes"></param>
        /// <returns></returns>
        public static INode GetFirstNode(this INode node, params Type[] excludedTypes)
        {
            if (!excludedTypes.Contains(node.GetType()))
            {
                return node;
            }

            return GetFirstNode(node.Children[0], excludedTypes);
        }

        public static IEnumerable<IdentifierReference> GetUnusedAssignmentIdentifierReferences(this INode node)
        {
            var nodes = new List<(AssignmentNode, IdentifierReference)>();

            var blockNodes = node.GetNodes(typeof(BlockNode));
            INode lastNode = default;
            foreach (var block in blockNodes)
            {
                var flattened = block.GetFlattenedNodes(typeof(GenericNode), typeof(BlockNode))
                    .OrderBy(n => n.Reference?.Selection.StartLine ?? n.SortOrder)
                    .ThenByDescending(n => n.Reference?.Selection.StartColumn)
                    .ToArray(); // INode.SortOrder puts LHS first
                foreach (var flattenedNode in flattened)
                {
                    if (flattenedNode is AssignmentNode && lastNode is AssignmentNode assignmentNode)
                    {
                        nodes.Add((assignmentNode, assignmentNode.Reference));
                    }

                    lastNode = flattenedNode;
                }

                var firstNonGenericNode = block.Children[0].GetFirstNode(typeof(GenericNode));
                if (lastNode is AssignmentNode node1 && firstNonGenericNode is DeclarationNode)
                {
                    nodes.Add((node1, node1.Reference));
                }
            }

            return nodes
                .Where(n => !n.Item1.Usages.Any())
                .Select(n => n.Item2);
        }
    }
}
