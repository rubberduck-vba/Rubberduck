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
        public LabelNode(string scope, Match match, string comment)
            : base(scope, match, comment)
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
