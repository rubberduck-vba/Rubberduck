using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Rubberduck.VBA.Parser.Grammar;

namespace Rubberduck.VBA.Parser
{
    [ComVisible(false)]
    public class ForLoopNode : CodeBlockNode
    {
        public ForLoopNode(Instruction instruction, string scope, Match match, IEnumerable<SyntaxTreeNode> nodes) 
            : base(instruction, scope, match, new[] {ReservedKeywords.Next}, null, nodes)
        {
        }

        public Expression Lower { get { return new Expression(RegexMatch.Groups["lower"].Value); } }
        public Expression Upper { get { return new Expression(RegexMatch.Groups["upper"].Value); } }
        public Expression Step { get { return new Expression(RegexMatch.Groups["step"].Value); } }
    }
}