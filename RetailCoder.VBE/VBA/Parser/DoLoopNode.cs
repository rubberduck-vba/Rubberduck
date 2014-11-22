using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Rubberduck.VBA.Parser.Grammar;

namespace Rubberduck.VBA.Parser
{
    [ComVisible(false)]
    public class DoLoopNode : CodeBlockNode
    {
        public DoLoopNode(Instruction instruction, string scope, Match match, IEnumerable<SyntaxTreeNode> nodes)
            : base(instruction, 
                   scope, 
                   match, 
                   new[] { 
                            ReservedKeywords.Loop, 
                            ReservedKeywords.Until, 
                            ReservedKeywords.While, 
                            ReservedKeywords.Wend 
                         }, 
                    null, nodes)
        {
        }

        /// <summary>
        /// Gets a <c>Boolean</c> <see cref="Expression"/>.
        /// </summary>
        public Expression Condition { get { return new Expression(RegexMatch.Groups["expression"].Value); } }
    }
}