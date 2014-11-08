using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RetailCoderVBE.Reflection.VBA
{
    internal class LogicalCodeLineNode : SyntaxTreeNode
    {
        public LogicalCodeLineNode(string scope, int startLine, string instructions)
            : this(scope, startLine, instructions, string.Empty)
        {

        }

        public LogicalCodeLineNode(string scope, int startLine, string instructions, string comment)
            : base(scope, comment)
        {
            _startLine = startLine;
        }

        private readonly int _startLine;
        private int StartLine
        {
            get
            {
                return _startLine;
            }
        }
    }
}
