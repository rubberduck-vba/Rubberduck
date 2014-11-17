using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Rubberduck.VBA.Parser
{
    [ComVisible(false)]
    public class CodeBlockNode : SyntaxTreeNode
    {
        public CodeBlockNode(Instruction instruction, string scope, Match match, string[] endingMarkers, Type childSyntaxType, IEnumerable<SyntaxTreeNode> nodes)
            : base(instruction, scope, match, nodes)
        {
            _endingMarkers = endingMarkers;
            _childSyntaxType = childSyntaxType;
        }

        private readonly string[] _endingMarkers;
        public IEnumerable<string> EndOfBlockMarkers { get { return _endingMarkers; } }

        private readonly Type _childSyntaxType;
        public Type ChildSyntaxType { get { return _childSyntaxType; } }

        /// <summary>
        /// Returns a new <see cref="CodeBlockNode"/> with the specified node appended to its child nodes collection.
        /// </summary>
        /// <param name="node"></param>
        /// <returns></returns>
        public virtual TNode AddNode<TNode>(SyntaxTreeNode node) where TNode : CodeBlockNode
        {
            return new CodeBlockNode(node.Instruction, Scope, RegexMatch, _endingMarkers, _childSyntaxType, ChildNodes.Concat(new[] { node })) as TNode;
        }
    }
}
