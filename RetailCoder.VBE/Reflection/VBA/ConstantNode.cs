using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Rubberduck.Reflection.VBA
{
    // todo: handle multiple declarations on single instruction.

    internal class ConstantNode : DeclarationNode
    {
        public ConstantNode(Instruction instruction, string scope, Match match)
            : base(instruction, scope, match)
        { }

        /// <summary>
        /// Gets the constant's value. Strings include delimiting quotes.
        /// </summary>
        public string Value { get { return RegexMatch.Groups["value"].Value; } }
    }
}
