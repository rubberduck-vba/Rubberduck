using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.VBEditor;

namespace Rubberduck.VBA.Nodes
{
    /// <summary>
    /// The base class for all nodes in a Rubberduck code tree. todo: delete this obsolete class.
    /// </summary>
    public abstract class Node : INode
    {
        private readonly ParserRuleContext _context;
        private readonly Selection _selection;
        private readonly string _parentScope;

        private readonly string _localScope;
        private readonly ICollection<Node> _childNodes;

        /// <summary>
        /// Represents a context in the code tree.
        /// </summary>
        /// <param name="context">The parser rule context, obtained from an ANTLR-generated parser method.</param>
        /// <param name="parentScope">The scope this context belongs to. <c>null</c> for the root context.</param>
        /// <param name="localScope">The scope this context defines, if any. <c>null</c> if omitted.</param>
        /// <param name="childNodes">The child nodes.</param>
        /// <remarks>
        /// Specifying a <c>localScope</c> ensures child nodes can be added, regardless of 
        /// </remarks>
        protected Node(ParserRuleContext context, string parentScope, string localScope = null, ICollection<Node> childNodes = null)
        {
            _context = context;
            _selection = context.GetSelection();
            _parentScope = parentScope;

            _localScope = localScope;

            _childNodes = (localScope != null && childNodes == null)
                            ? new List<Node>()
                            : childNodes;
        }

        /// <summary>
        /// Gets the parser rule context for the context.
        /// </summary>
        protected ParserRuleContext Context { get { return _context; } }

        /// <summary>
        /// Gets the name of the scope this context belongs to.
        /// </summary>
        public string ParentScope { get { return _parentScope; } }

        /// <summary>
        /// Gets the name of the scope defined by this context. <c>null</c> if context cannot be a parent.
        /// </summary>
        public string LocalScope { get { return _localScope; } }

        /// <summary>
        /// Gets a value representing the position of the context in the code Module.
        /// </summary>
        public Selection Selection { get { return _selection; } }
        
        /// <summary>
        /// Gets a the child nodes. <c>null</c> if context cannot be a parent.
        /// </summary>
        public IEnumerable<Node> Children { get { return _childNodes; } }

        /// <summary>
        /// Adds a child context.
        /// </summary>
        /// <param name="node">The child context to be added.</param>
        /// <exception cref="InvalidOperationException">Thrown if context cannot have child nodes.</exception>
        public void AddChild(Node node)
        {
            if (_childNodes == null)
            {
                throw new InvalidOperationException("This context cannot have child nodes.");
            }

            _childNodes.Add(node);
        }
    }
}