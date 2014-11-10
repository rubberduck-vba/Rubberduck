using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RetailCoderVBE.Reflection.VBA
{
    internal class ProcedureNode : SyntaxTreeNode
    {
        public ProcedureNode(string scope, Match match, string comment)
            : base(scope, match, comment, true)
        {

        }
    }
}
