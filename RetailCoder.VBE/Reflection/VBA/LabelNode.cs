using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Rubberduck.Reflection.VBA
{
    internal class LabelNode : SyntaxTreeNode
    {
        public LabelNode(Instruction instruction, string scope, Match match)
            : base(instruction, scope, match)
        {
        }

        public string Label
        { 
            get 
            {
                return RegexMatch.Groups["identifier"].Value;
            } 
        }
    }
}
