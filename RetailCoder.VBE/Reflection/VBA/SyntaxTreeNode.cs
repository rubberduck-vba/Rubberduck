using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Rubberduck.Reflection.VBA
{
    internal abstract class SyntaxTreeNode
    {
        public SyntaxTreeNode(string scope)
            : this(scope, null, string.Empty) 
        {
        }

        public SyntaxTreeNode(string scope, string comment)
            : this(scope, null, comment)
        {
        }

        public SyntaxTreeNode(string scope, Match match)
            : this(scope, match, string.Empty)
        {

        }

        public SyntaxTreeNode(string scope, Match match, bool hasChildNodes)
            : this(scope, match, string.Empty, hasChildNodes)
        {

        }

        public SyntaxTreeNode(string scope, Match match, string comment, bool hasChildNodes = false)
        {
            _scope = scope;
            _match = match;
            _comment = comment;
            _hasChildNodes = hasChildNodes;
        }

        private readonly string _scope;
        public string Scope { get { return _scope; } }

        private readonly bool _hasChildNodes;
        public bool HasChildNodes { get { return _hasChildNodes; } }

        private readonly Match _match;
        protected Match RegexMatch { get { return _match; } }

        private readonly string _comment;
        public string Comment { get { return _comment; } }
    }
}
