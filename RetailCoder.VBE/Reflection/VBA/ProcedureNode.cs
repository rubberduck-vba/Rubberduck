using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Rubberduck.Reflection.VBA
{
    internal class ProcedureNode : CodeBlockNode
    {
        public ProcedureNode(Instruction instruction, string scope, Match match, string keyword)
            : base(instruction, scope, match, ReservedKeywords.End + " " + keyword, typeof(SyntaxTreeNode))
        {

        }
    }
}
