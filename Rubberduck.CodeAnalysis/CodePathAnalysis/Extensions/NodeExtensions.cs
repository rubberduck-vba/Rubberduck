using Rubberduck.Inspections.CodePathAnalysis.Nodes;
using System;
using System.Collections.Generic;
using System.Linq;

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

        public static IEnumerable<AssignmentNode> GetAssignmentNodes(this INode node)
        {
            var nodes = new List<AssignmentNode>();

            var blocks = node.GetNodes(typeof(BlockNode));
            AssignmentNode previous = default;
            foreach (var block in blocks)
            {
                var flattened = block.GetFlattenedNodes(typeof(GenericNode), typeof(BlockNode));
                foreach (var current in flattened)
                {
                    switch (current)
                    {
                        case AssignmentNode assignmentNode:
                            previous = assignmentNode;
                            nodes.Add(previous);
                            break;
                        case ReferenceNode referenceNode:
                            previous?.AddUsage(referenceNode);
                            break;
                    }
                }
            }

            return nodes;
        }
    }
}
