using Rubberduck.Inspections.CodePathAnalysis.Nodes;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections.CodePathAnalysis.Extensions
{
    public static class NodeExtensions
    {
        public static IEnumerable<INode> FlattenedNodes(this INode node, IEnumerable<Type> excludedTypes)
        {
            foreach (var child in node.Children)
            {
                if (!excludedTypes.Contains(child.GetType()))
                {
                    yield return child;
                }
                else
                {
                    foreach (var nextChild in FlattenedNodes(child, excludedTypes))
                    {
                        yield return nextChild;
                    }
                }
            }
        }

        public static IEnumerable<INode> Nodes(this INode node, ICollection<Type> types)
        {
            if (types.Contains(node.GetType()))
            {
                yield return node;
            }

            foreach (var child in node.Children)
            {
                foreach (var childNode in Nodes(child, types))
                {
                    yield return childNode;
                }
            }
        }

        public static INode GetFirstNode(this INode node, ICollection<Type> excludedTypes)
        {
            if (!excludedTypes.Contains(node.GetType()))
            {
                return node;
            }

            if (!node.Children.Any())
            {
                return null;
            }

            return GetFirstNode(node.Children[0], excludedTypes);
        }
    }
}
