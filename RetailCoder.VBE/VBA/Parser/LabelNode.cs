using System.Text.RegularExpressions;

namespace Rubberduck.VBA.Parser
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
