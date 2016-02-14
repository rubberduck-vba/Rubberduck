using System;
using System.Collections.Generic;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.Nodes
{
    // todo: delete this obsolete interface
    public interface INode
    {
        /// <summary>
        /// Gets the name of the scope this context belongs to.
        /// </summary>
        string ParentScope { get; }

        /// <summary>
        /// Gets the name of the scope defined by this context. <c>null</c> if context cannot be a parent.
        /// </summary>
        string LocalScope { get; }

        /// <summary>
        /// Gets a value representing the position of the context in the code Module.
        /// </summary>
        Selection Selection { get; }

        /// <summary>
        /// Gets a the child nodes. <c>null</c> if context cannot be a parent.
        /// </summary>
        IEnumerable<Node> Children { get; }

        /// <summary>
        /// Adds a child context.
        /// </summary>
        /// <param name="node">The child context to be added.</param>
        /// <exception cref="InvalidOperationException">Thrown if context cannot have child nodes.</exception>
        void AddChild(Node node);
    }
}