using Rubberduck.Inspections.CodePathAnalysis.Nodes;
using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.CodePathAnalysis.Extensions
{
    public static class NodeExtensions
    {
        public static IEnumerable<INode> GetFlattenedNodes(this INode node, IEnumerable<Type> excludedTypes)
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

        public static IEnumerable<INode> GetNodes(this INode node, IEnumerable<Type> types)
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

        public static INode GetFirstNode(this INode node, IEnumerable<Type> excludedTypes)
        {
            if (!excludedTypes.Contains(node.GetType()))
            {
                return node;
            }

            return GetFirstNode(node.Children[0], excludedTypes);
        }

        public static List<IdentifierReference> GetIdentifierReferences(this INode node)
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
