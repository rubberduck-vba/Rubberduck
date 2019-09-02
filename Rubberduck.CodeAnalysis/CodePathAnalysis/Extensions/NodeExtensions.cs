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
                        case BranchNode branchNode:
                            //previous = default;
                            break;
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
