using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RetailCoderVBE.Reflection.VBA
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

        public SyntaxTreeNode(string scope, Match match, bool definesScope)
            : this(scope, match, string.Empty, definesScope)
        {

        }

        public SyntaxTreeNode(string scope, Match match, string comment, bool definesScope = false)
        {
            _scope = scope;
            _match = match;
            _comment = comment;
            _definesScope = definesScope;
        }

        private readonly string _scope;
        public string Scope { get { return _scope; } }

        private readonly bool _definesScope;
        public bool DefinesScope { get { return _definesScope; } }

        private readonly Match _match;
        public Match RegexMatch { get { return _match; } }

        private readonly string _comment;
        public string Comment { get { return _comment; } }
    }
}
