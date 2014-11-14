using System.Text.RegularExpressions;

namespace Rubberduck.Reflection.VBA
{
    internal abstract class SyntaxTreeNode
    {
        protected SyntaxTreeNode(Instruction instruction, string scope)
            : this(instruction, scope, null)
        {
        }

        protected SyntaxTreeNode(Instruction instruction, string scope, Match match, bool hasChildNodes = false)
        {
            _instruction = instruction;
            _scope = scope;
            _match = match;
            _hasChildNodes = hasChildNodes;
        }

        private readonly Instruction _instruction;
        public Instruction Instruction { get { return _instruction; } }

        private readonly string _scope;
        public string Scope { get { return _scope; } }

        private readonly bool _hasChildNodes;
        public bool HasChildNodes { get { return _hasChildNodes; } }

        private readonly Match _match;
        protected Match RegexMatch { get { return _match; } }
    }
}
