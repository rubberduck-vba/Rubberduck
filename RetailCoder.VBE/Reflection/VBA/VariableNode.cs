using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Rubberduck.Reflection.VBA
{
    // todo: handle multiple declarations on single instruction - grammar/regex already supports it.

    internal class VariableNode : DeclarationNode
    {
        public VariableNode(Instruction instruction, string scope, Match match)
            : base(instruction, scope, match)
        { }

    }
}
