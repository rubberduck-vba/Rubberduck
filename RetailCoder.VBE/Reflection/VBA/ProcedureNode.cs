using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Rubberduck.Reflection.VBA
{
    internal class ProcedureNode : SyntaxTreeNode
    {
        public ProcedureNode(Instruction instruction, string scope, Match match)
            : base(instruction, scope, match, true)
        {

        }
    }
}
