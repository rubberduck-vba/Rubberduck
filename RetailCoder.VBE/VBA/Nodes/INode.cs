using System;
using System.Collections.Generic;
using Rubberduck.Extensions;

namespace Rubberduck.VBA.Nodes
{
    public interface INode
    {
        /// <summary>
        /// Gets the name of the scope this node belongs to.
        /// </summary>
        string ParentScope { get; }

        /// <summary>
        /// Gets the name of the scope defined by this node. <c>null</c> if node cannot be a parent.
        /// </summary>
        string LocalScope { get; }

        /// <summary>
        /// Gets a value representing the position of the node in the code module.
        /// </summary>
        Selection Selection { get; }

        /// <summary>
        /// Gets a the child nodes. <c>null</c> if node cannot be a parent.
        /// </summary>
        IEnumerable<Node> Children { get; }

        /// <summary>
        /// Adds a child node.
        /// </summary>
        /// <param name="node">The child node to be added.</param>
        /// <exception cref="InvalidOperationException">Thrown if node cannot have child nodes.</exception>
        void AddChild(Node node);
    }
}