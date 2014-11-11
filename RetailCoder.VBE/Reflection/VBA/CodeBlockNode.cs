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
        public CodeBlockNode(string scope, Match match, string comment, string endingMarker, Type childNodeType)
            : this(scope, match, comment, endingMarker, childNodeType, new List<SyntaxTreeNode>())
        {

        }

        private CodeBlockNode(string scope, Match match, string comment, string endingMarker, Type childSyntaxType, IEnumerable<SyntaxTreeNode> nodes)
            : base(scope, match, comment, true)
        {
            _endingMarker = endingMarker;
            _childSyntaxType = childSyntaxType;
            _nodes = nodes;
        }

        private readonly string _endingMarker;
        public string EndOfBlockMarker { get { return _endingMarker; } }

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
            return new CodeBlockNode(Scope, RegexMatch, Comment, _endingMarker, _childSyntaxType, _nodes.Concat(new[] { node }));
        }
    }
}
