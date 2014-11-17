using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Rubberduck.VBA.Parser
{
    [ComVisible(false)]
    public abstract class SyntaxTreeNode
    {
        protected SyntaxTreeNode(Instruction instruction, string scope, Match match = null, IEnumerable<SyntaxTreeNode> childNodes = null)
        {
            _instruction = instruction;
            _scope = scope;
            _match = match;
            _childNodes = childNodes;
        }

        private readonly Instruction _instruction;
        public Instruction Instruction { get { return _instruction; } }

        private readonly string _scope;
        public string Scope { get { return _scope; } }

        private readonly IEnumerable<SyntaxTreeNode> _childNodes;
        public IEnumerable<SyntaxTreeNode> ChildNodes { get { return _childNodes; } }

        public bool HasChildNodes { get { return _childNodes != null; } }

        private readonly Match _match;
        protected Match RegexMatch { get { return _match; } }
    }
}
