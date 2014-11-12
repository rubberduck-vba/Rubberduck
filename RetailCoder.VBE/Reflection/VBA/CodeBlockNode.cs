using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Rubberduck.Reflection.VBA
{
    internal class CodeBlockNode : SyntaxTreeNode
    {
        public CodeBlockNode(Instruction instruction, string scope, Match match, string endingMarker, IEnumerable<SyntaxTreeNode> childNodes)
            : this(instruction, scope, match, new[] { endingMarker }, typeof(SyntaxTreeNode), childNodes)
        {

        }

        public CodeBlockNode(Instruction instruction, string scope, Match match, string endingMarker, Type childNodeType)
            : this(instruction, scope, match, new[] {endingMarker}, childNodeType, new List<SyntaxTreeNode>())
        {

        }

        public CodeBlockNode(Instruction instruction, string scope, Match match, string[] endingMarkers, Type childNodeType)
            : this(instruction, scope, match, endingMarkers, childNodeType, new List<SyntaxTreeNode>())
        {

        }

        private CodeBlockNode(Instruction instruction, string scope, Match match, string[] endingMarkers, Type childSyntaxType, IEnumerable<SyntaxTreeNode> nodes)
            : base(instruction, scope, match, true)
        {
            _endingMarkers = endingMarkers;
            _childSyntaxType = childSyntaxType;
            _nodes = nodes;
        }

        private readonly string[] _endingMarkers;
        public string[] EndOfBlockMarkers { get { return _endingMarkers; } }

        private readonly IEnumerable<SyntaxTreeNode> _nodes;
        public IEnumerable<SyntaxTreeNode> ChildNodes { get { return _nodes; } }

        private readonly Type _childSyntaxType;
        public Type ChildSyntaxType { get { return _childSyntaxType; } }

        /// <summary>
        /// Returns a new <see cref="CodeBlockNode"/> with the specified node appended to its child nodes collection.
        /// </summary>
        /// <param name="node"></param>
        /// <returns></returns>
        public CodeBlockNode AddNode(SyntaxTreeNode node)
        {
            return new CodeBlockNode(node.Instruction, Scope, RegexMatch, _endingMarkers, _childSyntaxType, _nodes.Concat(new[] { node }));
        }
    }
}
