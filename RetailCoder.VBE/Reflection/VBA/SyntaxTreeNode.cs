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
        public SyntaxTreeNode(Match match)
        {
            _match = match;
        }

        private readonly Match _match;
        public Match RegexMatch { get { return _match; } }
    }
}
