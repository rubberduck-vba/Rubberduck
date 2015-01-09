using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA
{
    [ComVisible(false)]
    public class IfBlockNode : CodeBlockNode
    {
        public IfBlockNode(Instruction instruction, string scope, Match match, IEnumerable<SyntaxTreeNode> nodes)
            : base(instruction, scope, match, match.Groups["expression"].Success ? new string[]{} : new[] {ReservedKeywords.End+" "+ReservedKeywords.If, ReservedKeywords.Else, ReservedKeywords.ElseIf}, null, nodes)
        {
        }

        /// <summary>
        /// Gets a <c>Boolean</c> <see cref="Expression"/>.
        /// </summary>
        public Expression Condition { get { return new Expression(RegexMatch.Groups["condition"].Value); } }

        /// <summary>
        /// Gets an <see cref="Expression"/>
        /// </summary>
        public Expression Expression { get { return new Expression(RegexMatch.Groups["expression"].Value); } }
    }
}
